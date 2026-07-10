import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { AuditService } from '../audit/AuditService';
import type { NamingPolicyService } from '../namingPolicy/NamingPolicyService';
import type { UserService } from '../users/UserService';
import type { SiteAccessService } from '../sites/SiteAccessService';
import type { TelemetryService } from '../telemetry/TelemetryService';
import { GraphServiceError, RequestAbortedError } from '../graph/GraphError';
import { generateTempPassword } from '../passwords/passwordGenerator';
import type { IAppSettings, IJobStep, IProvisioningJob, JobType } from '../../models';
import { DEFAULT_APP_SETTINGS, isOnboardingPayload } from '../../models';
import { assertTransition, canStartJob } from './jobStateMachine';
import { stepsForJobType } from './stepRegistry';
import { StepFailure } from './stepTypes';
import type {
  ICredentialPresentation,
  IJobSecrets,
  IStepContext,
  IWorkflowStepDefinition
} from './stepTypes';
import { delay, targetUserOf } from './stepHelpers';

export interface IEngineDependencies {
  graph: GraphService;
  data: SharePointDataService;
  audit: AuditService;
  naming: NamingPolicyService;
  users: UserService;
  siteAccess: SiteAccessService;
  /** Optional tenant settings; defaults apply when absent (engine-only callers in tests). */
  settings?: IAppSettings;
  /** Optional telemetry sink; null-safe when omitted. */
  telemetry?: TelemetryService;
}

export interface IEngineCallbacks {
  /** Fired after every persisted state change — drives the live progress UI. */
  onJobUpdated?: (job: IProvisioningJob) => void;
  /** Copy-once credential hand-off; resolves when the operator confirms. */
  presentCredentials?: (credential: ICredentialPresentation) => Promise<void>;
}

export interface IEngineOptions {
  /** Base for exponential backoff between step attempts (test override). */
  stepBackoffBaseMs?: number;
}

/**
 * Client-side provisioning engine (spec Section 6). Steps execute
 * sequentially in the operator's browser session; each is idempotent and
 * state is persisted to StepsJson after every attempt, so a closed tab
 * resumes cleanly from the first non-completed step.
 */
export class WorkflowEngine {
  private readonly _deps: IEngineDependencies;
  private readonly _backoffBaseMs: number;
  private readonly _runningItems: Set<number> = new Set();
  /**
   * Steps flagged for skip while their job's _execute() loop is already
   * in-flight. SPFx's Graph client gives no real request cancellation (only
   * cooperative AbortSignal checks between attempts), so a step's in-flight
   * call can still resolve successfully after skipStep() has already
   * persisted 'skipped' for it. _runStepWithRetries consults this before
   * persisting 'completed'/'failed' so that write can't clobber the skip.
   */
  private readonly _pendingSkips: Map<number, Set<string>> = new Map();
  private _settings: IAppSettings;

  public constructor(deps: IEngineDependencies, options?: IEngineOptions) {
    this._deps = deps;
    this._backoffBaseMs = options?.stepBackoffBaseMs ?? 1000;
    this._settings = deps.settings ?? DEFAULT_APP_SETTINGS;
  }

  /** Update tenant settings at runtime (SettingsContext calls this on load/change). */
  public updateSettings(settings: IAppSettings): void {
    this._settings = settings;
  }

  /** Initial StepsJson content for a new job of the given type. */
  public buildInitialSteps(jobType: JobType = 'Onboard'): IJobStep[] {
    return stepsForJobType(jobType).map((definition) => ({
      stepId: definition.id,
      status: 'pending',
      attempts: 0,
      maxAttempts: definition.maxAttempts,
      lastError: null,
      startedUtc: null,
      completedUtc: null,
      skippable: definition.skippable
    }));
  }

  public isRunning(itemId: number): boolean {
    return this._runningItems.has(itemId);
  }

  /** Start or resume a job. Refuses anything the state machine says is not startable. */
  public async runJob(
    itemId: number,
    callbacks?: IEngineCallbacks,
    signal?: AbortSignal
  ): Promise<IProvisioningJob> {
    if (this._runningItems.has(itemId)) {
      throw new Error(`Job ${itemId} is already running in this session`);
    }
    this._runningItems.add(itemId);
    const started: number = Date.now();
    try {
      const job: IProvisioningJob = await this._execute(itemId, callbacks, signal);
      if (this._deps.telemetry) {
        this._deps.telemetry.trackEvent(
          'engine.runJob.complete',
          {
            jobId: job.jobId,
            jobType: job.jobType,
            status: job.status,
            durationMs: Date.now() - started
          },
          job.status === 'Completed' ? 'info' : 'warning'
        );
      }
      return job;
    } catch (err) {
      if (this._deps.telemetry && !(err instanceof RequestAbortedError)) {
        this._deps.telemetry.trackError(err, {
          jobId: String(itemId),
          durationMs: Date.now() - started
        });
      }
      throw err;
    } finally {
      this._runningItems.delete(itemId);
      this._pendingSkips.delete(itemId);
    }
  }

  /** Manual retry: reset the failed step and resume the job. */
  public async retryStep(
    itemId: number,
    stepId: string,
    callbacks?: IEngineCallbacks,
    signal?: AbortSignal
  ): Promise<IProvisioningJob> {
    if (this._runningItems.has(itemId)) {
      // A failed, non-continuable step blocks the pipeline before the loop
      // reaches later steps, but continueOnFailure steps let it run past a
      // failure — so a failed step's job can still be actively executing.
      // Reject outright rather than writing StepsJson underneath the
      // in-flight run's own full-array persists (no ETag/optimistic
      // concurrency exists on that write to arbitrate a real race).
      throw new Error(`Job ${itemId} is currently running — wait for it to finish before retrying a step`);
    }
    const job: IProvisioningJob = await this._deps.data.getJob(itemId);
    const step: IJobStep | undefined = job.steps.filter((s) => s.stepId === stepId)[0];
    if (!step || step.status !== 'failed') {
      throw new Error(`Step ${stepId} is not in a retryable state`);
    }
    step.status = 'pending';
    step.attempts = 0;
    step.lastError = null;
    step.startedUtc = null;
    step.completedUtc = null;
    await this._deps.data.updateJobSteps(itemId, job.steps);
    return this.runJob(itemId, callbacks, signal);
  }

  /** Skip a failed or running step, when its definition allows it, then resume. */
  public async skipStep(
    itemId: number,
    stepId: string,
    callbacks?: IEngineCallbacks,
    signal?: AbortSignal
  ): Promise<IProvisioningJob> {
    const job: IProvisioningJob = await this._deps.data.getJob(itemId);
    const step: IJobStep | undefined = job.steps.filter((s) => s.stepId === stepId)[0];
    if (!step || !step.skippable) {
      throw new Error(`Step ${stepId} is not skippable`);
    }
    // Flag the skip BEFORE writing it, while the job is still running: if
    // the in-flight step's Graph call is already in flight (the common
    // case — see the class comment on _pendingSkips), _runStepWithRetries
    // checks this flag before it would otherwise persist 'completed'/
    // 'failed' over the top of this skip.
    const runningNow: boolean = this._runningItems.has(itemId);
    if (runningNow) {
      const pending: Set<string> = this._pendingSkips.get(itemId) ?? new Set<string>();
      pending.add(stepId);
      this._pendingSkips.set(itemId, pending);
    }
    step.status = 'skipped';
    step.completedUtc = new Date().toISOString();
    await this._deps.data.updateJobSteps(itemId, job.steps);
    await this._deps.audit.log({
      jobId: job.jobId,
      action: `skip-step:${stepId}`,
      targetUser: targetUserOf(job.payload),
      graphEndpoint: '',
      requestSummary: '{}',
      responseCode: 0,
      durationMs: 0,
      result: 'Skipped',
      correlationId: job.correlationId
    });
    // If the job is already running in this session (the operator clicked
    // Skip on a running step), the in-flight pass owns persisting further
    // progress — don't start a second concurrent run.
    if (runningNow) {
      return job;
    }
    return this.runJob(itemId, callbacks, signal);
  }

  public async cancelJob(itemId: number): Promise<void> {
    const job: IProvisioningJob = await this._deps.data.getJob(itemId);
    assertTransition(job.status, 'Cancelled');
    await this._deps.data.updateJobStatus(itemId, 'Cancelled');
    await this._deps.audit.log({
      jobId: job.jobId,
      action: 'cancel-job',
      targetUser: targetUserOf(job.payload),
      graphEndpoint: '',
      requestSummary: '{}',
      responseCode: 0,
      durationMs: 0,
      result: 'Success',
      correlationId: job.correlationId
    });
  }

  /**
   * Re-issues the target user's temporary password (or TAP) for a completed
   * onboarding/clone job — e.g. the operator lost the original one-time
   * hand-off. Not a pipeline step: it doesn't touch StepsJson or job status,
   * only writes a fresh credential to Graph and hands it to the caller once.
   */
  public async regenerateCredentials(
    itemId: number,
    callbacks?: IEngineCallbacks,
    signal?: AbortSignal
  ): Promise<void> {
    const job: IProvisioningJob = await this._deps.data.getJob(itemId);
    if (job.status !== 'Completed' && job.status !== 'PartiallyFailed') {
      throw new Error('Credentials can only be regenerated for a completed job');
    }
    if (job.jobType !== 'Onboard' && job.jobType !== 'Clone') {
      throw new Error('Credential regeneration only applies to onboarding and clone jobs');
    }
    if (!isOnboardingPayload(job.payload)) {
      throw new Error('Job payload is not an onboarding payload');
    }
    const payload = job.payload;
    if (payload.identity.accountType === 'guest') {
      throw new Error('Guest invitations have no password or TAP to regenerate');
    }
    if (!job.targetUserId) {
      throw new Error('This job has no target user');
    }
    const userId: string = job.targetUserId;
    const upn: string = payload.identity.userPrincipalName;
    const mode = payload.accountSettings.credentialMode;
    const started: number = Date.now();
    const endpoint: string =
      mode === 'tap' ? `/users/${userId}/authentication/temporaryAccessPassMethods` : `/users/${userId}`;

    try {
      let value: string;
      if (mode === 'tap') {
        const existing: { value: { id: string }[] } = await this._deps.graph.get<{
          value: { id: string }[];
        }>(endpoint, { signal });
        for (const method of existing.value ?? []) {
          await this._deps.graph.delete<void>(`${endpoint}/${method.id}`, { signal });
        }
        const created: { temporaryAccessPass: string } = await this._deps.graph.post<{
          temporaryAccessPass: string;
        }>(endpoint, {}, { signal });
        value = created.temporaryAccessPass;
      } else {
        value = generateTempPassword();
        const body = { passwordProfile: { password: value, forceChangePasswordNextSignIn: true } };
        await this._deps.graph.patch<void>(endpoint, body, { signal });
      }
      await this._deps.audit.log({
        jobId: job.jobId,
        action: 'regenerate-credentials',
        targetUser: upn,
        graphEndpoint: endpoint,
        requestSummary: JSON.stringify({ mode }),
        responseCode: 200,
        durationMs: Date.now() - started,
        result: 'Success',
        correlationId: job.correlationId
      });
      await callbacks?.presentCredentials?.({ kind: mode, value, userPrincipalName: upn });
    } catch (err) {
      if (err instanceof RequestAbortedError) {
        throw err;
      }
      const graphError: GraphServiceError =
        err instanceof GraphServiceError
          ? err
          : new GraphServiceError((err as Error)?.message ?? 'Unknown error', 0, 'UnknownError', '');
      await this._deps.audit.log({
        jobId: job.jobId,
        action: 'regenerate-credentials',
        targetUser: upn,
        graphEndpoint: endpoint,
        requestSummary: JSON.stringify({ mode }),
        responseCode: graphError.statusCode,
        durationMs: Date.now() - started,
        result: 'Failure',
        correlationId: job.correlationId
      });
      throw graphError;
    }
  }

  private async _execute(
    itemId: number,
    callbacks?: IEngineCallbacks,
    signal?: AbortSignal
  ): Promise<IProvisioningJob> {
    const job: IProvisioningJob = await this._deps.data.getJob(itemId);
    if (!canStartJob(job.status)) {
      throw new Error(
        `Job ${job.jobId} cannot start from status ${job.status} — approval is required first`
      );
    }
    if (job.status !== 'Running') {
      assertTransition(job.status, 'Running');
      await this._deps.data.updateJobStatus(itemId, 'Running');
      job.status = 'Running';
      callbacks?.onJobUpdated?.(job);
    }

    // Merge persisted step state with the registry (registry order wins; new
    // steps added by an upgrade appear as pending).
    const pipeline = stepsForJobType(job.jobType);
    const persisted: Map<string, IJobStep> = new Map(job.steps.map((s) => [s.stepId, s]));
    job.steps = pipeline.map(
      (definition) =>
        persisted.get(definition.id) ?? {
          stepId: definition.id,
          status: 'pending',
          attempts: 0,
          maxAttempts: definition.maxAttempts,
          lastError: null,
          startedUtc: null,
          completedUtc: null,
          skippable: definition.skippable
        }
    );

    const secrets: IJobSecrets = {};
    const ctx: IStepContext = {
      graph: this._deps.graph,
      data: this._deps.data,
      audit: this._deps.audit,
      naming: this._deps.naming,
      users: this._deps.users,
      siteAccess: this._deps.siteAccess,
      job,
      secrets,
      signal,
      settings: this._settings,
      presentCredentials: async (credential: ICredentialPresentation): Promise<void> => {
        if (callbacks?.presentCredentials) {
          await callbacks.presentCredentials(credential);
        }
      }
    };

    let blocked: boolean = false;
    for (const definition of pipeline) {
      const state: IJobStep = job.steps.filter((s) => s.stepId === definition.id)[0];
      if (state.status === 'completed' || state.status === 'skipped') {
        continue;
      }
      if (state.status === 'failed' && state.attempts >= state.maxAttempts) {
        // Exhausted in an earlier session; manual retry/skip is the only way on.
        if (!definition.continueOnFailure) {
          blocked = true;
          break;
        }
        continue;
      }

      try {
        await this._runStepWithRetries(definition, state, ctx, itemId, callbacks, signal);
      } catch (err) {
        if (err instanceof RequestAbortedError) {
          // Leave the job Running; the persisted StepsJson resumes it later.
          return job;
        }
        throw err;
      }

      if (state.status === 'failed' && !definition.continueOnFailure) {
        blocked = true;
        break;
      }
    }

    const anyFailed: boolean = job.steps.some((s) => s.status === 'failed');
    const allDone: boolean = job.steps.every(
      (s) => s.status === 'completed' || s.status === 'skipped'
    );
    // When a non-continuable step fails and blocks the pipeline, the job is
    // Failed (no further progress possible without manual retry). When some
    // steps failed but the pipeline continued (continueOnFailure), the job
    // is PartiallyFailed. When all steps completed/skipped, it's Completed.
    const finalStatus: IProvisioningJob['status'] = allDone
      ? 'Completed'
      : blocked && !anyFailed
        ? 'PartiallyFailed'
        : blocked && anyFailed
          ? 'Failed'
          : anyFailed
            ? 'PartiallyFailed'
            : 'Completed';
    assertTransition(job.status, finalStatus);
    await this._deps.data.updateJobStatus(itemId, finalStatus);
    job.status = finalStatus;
    // Wipe in-memory secrets as soon as the run ends.
    delete secrets.temporaryPassword;
    delete secrets.temporaryAccessPass;
    callbacks?.onJobUpdated?.(job);
    return job;
  }

  private async _runStepWithRetries(
    definition: IWorkflowStepDefinition,
    state: IJobStep,
    ctx: IStepContext,
    itemId: number,
    callbacks?: IEngineCallbacks,
    signal?: AbortSignal
  ): Promise<void> {
    while (state.attempts < state.maxAttempts) {
      state.status = 'running';
      if (!state.startedUtc) {
        state.startedUtc = new Date().toISOString();
      }
      await this._persist(itemId, ctx.job, callbacks);

      try {
        await definition.run(ctx);
        if (this._consumePendingSkip(itemId, definition.id)) {
          // skipStep() already persisted 'skipped' for this step while this
          // call was in flight — its own write is authoritative; don't
          // clobber it by persisting 'completed' here too.
          state.status = 'skipped';
          state.completedUtc = state.completedUtc ?? new Date().toISOString();
          return;
        }
        state.status = 'completed';
        state.completedUtc = new Date().toISOString();
        state.lastError = null;
        await this._persist(itemId, ctx.job, callbacks);
        return;
      } catch (err) {
        if (err instanceof RequestAbortedError) {
          // The operator may have skipped this step while it was running.
          // Re-read the persisted state: if it was skipped, honor that;
          // otherwise reset to pending so a resume picks it up again.
          const refreshed: IProvisioningJob = await this._deps.data.getJob(itemId);
          const refreshedStep: IJobStep | undefined = refreshed.steps.filter(
            (s) => s.stepId === definition.id
          )[0];
          this._consumePendingSkip(itemId, definition.id);
          if (refreshedStep && refreshedStep.status === 'skipped') {
            state.status = 'skipped';
            state.completedUtc = refreshedStep.completedUtc;
          } else {
            state.status = 'pending';
          }
          await this._persist(itemId, ctx.job, callbacks);
          throw err;
        }
        if (this._consumePendingSkip(itemId, definition.id)) {
          state.status = 'skipped';
          state.completedUtc = state.completedUtc ?? new Date().toISOString();
          return;
        }
        const failure: StepFailure =
          err instanceof StepFailure
            ? err
            : new StepFailure((err as Error)?.message ?? 'Unknown error', 'UnknownError', false);
        state.attempts += 1;
        state.status = 'failed';
        state.lastError = {
          graphCode: failure.graphCode,
          message: failure.message,
          retryable: failure.retryable
        };
        await this._persist(itemId, ctx.job, callbacks);
        if (!failure.retryable || state.attempts >= state.maxAttempts) {
          return;
        }
        await delay(this._backoffBaseMs * Math.pow(2, state.attempts - 1), signal);
      }
    }
  }

  /** Returns true and clears the flag if stepId was skipped mid-flight for itemId. */
  private _consumePendingSkip(itemId: number, stepId: string): boolean {
    const pending: Set<string> | undefined = this._pendingSkips.get(itemId);
    if (!pending || !pending.has(stepId)) {
      return false;
    }
    pending.delete(stepId);
    if (pending.size === 0) {
      this._pendingSkips.delete(itemId);
    }
    return true;
  }

  private async _persist(
    itemId: number,
    job: IProvisioningJob,
    callbacks?: IEngineCallbacks
  ): Promise<void> {
    await this._deps.data.updateJobSteps(itemId, job.steps);
    callbacks?.onJobUpdated?.(job);
  }
}
