import { WorkflowEngine } from './WorkflowEngine';
import { GraphServiceError } from '../graph/GraphError';
import { AuditService } from '../audit/AuditService';
import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { NamingPolicyService } from '../namingPolicy/NamingPolicyService';
import type { UserService } from '../users/UserService';
import type { SiteAccessService } from '../sites/SiteAccessService';
import type {
  IAuditEntry,
  IJobStep,
  IOnboardingPayload,
  IProvisioningJob,
  JobStatus
} from '../../models';
import type { ICredentialPresentation } from './stepTypes';

// ---- test doubles -----------------------------------------------------------

type Handler = (path: string, body?: unknown) => unknown;

class MockGraph {
  public handlers: Record<string, Handler> = {};
  public calls: { method: string; path: string; body?: unknown }[] = [];

  private _dispatch(method: string, path: string, body?: unknown): unknown {
    this.calls.push({ method, path, body });
    for (const key of Object.keys(this.handlers)) {
      const [m, prefix] = key.split(' ');
      if (m === method && path.startsWith(prefix)) {
        return this.handlers[key](path, body);
      }
    }
    throw new GraphServiceError(`No handler for ${method} ${path}`, 404, 'Request_ResourceNotFound', 'test-req');
  }

  public get = async (path: string): Promise<unknown> => this._dispatch('GET', path);
  public post = async (path: string, body: unknown): Promise<unknown> => this._dispatch('POST', path, body);
  public patch = async (path: string, body: unknown): Promise<unknown> => this._dispatch('PATCH', path, body);
  public put = async (path: string, body: unknown): Promise<unknown> => this._dispatch('PUT', path, body);
  public delete = async (path: string): Promise<unknown> => this._dispatch('DELETE', path);
  public batch = async (): Promise<Map<string, unknown>> => new Map();

  public countCalls(method: string, pathPrefix: string): number {
    return this.calls.filter((c) => c.method === method && c.path.startsWith(pathPrefix)).length;
  }
}

class MockData {
  public status: JobStatus;
  public stepsJson: string;
  public targetUserId: string | null = null;
  public auditEntries: IAuditEntry[] = [];
  private readonly _payload: IOnboardingPayload;

  public constructor(status: JobStatus, payload: IOnboardingPayload, steps: IJobStep[]) {
    this.status = status;
    this._payload = payload;
    this.stepsJson = JSON.stringify(steps);
  }

  public getJob = async (itemId: number): Promise<IProvisioningJob> => ({
    itemId,
    jobId: 'job-guid-1',
    jobType: 'Onboard',
    status: this.status,
    payload: JSON.parse(JSON.stringify(this._payload)),
    steps: JSON.parse(this.stepsJson),
    scheduledFor: null,
    requestedBy: 'HR Person',
    approvedBy: null,
    correlationId: 'corr-1',
    targetUserId: this.targetUserId,
    createdUtc: '2026-01-01T00:00:00Z'
  });

  public updateJobStatus = async (_itemId: number, status: JobStatus): Promise<void> => {
    this.status = status;
  };

  public updateJobSteps = async (_itemId: number, steps: IJobStep[]): Promise<void> => {
    this.stepsJson = JSON.stringify(steps);
  };

  public setJobTargetUser = async (_itemId: number, userId: string): Promise<void> => {
    this.targetUserId = userId;
  };

  public addAuditEntry = async (entry: IAuditEntry): Promise<void> => {
    this.auditEntries.push(entry);
  };

  public steps(): IJobStep[] {
    return JSON.parse(this.stepsJson);
  }
}

function payload(): IOnboardingPayload {
  return {
    schemaVersion: 1,
    kind: 'onboard',
    personal: {
      firstName: 'Anna',
      lastName: 'Svensson',
      displayName: 'Anna Svensson',
      employeeId: 'E12345'
    },
    employment: {
      jobTitle: 'Controller',
      department: 'Finance',
      employeeType: 'Employee',
      hireDate: '2026-08-01',
      managerId: '11111111-2222-3333-4444-555555555555'
    },
    identity: {
      userPrincipalName: 'anna.svensson@contoso.com',
      mailNickname: 'anna.svensson',
      domain: 'contoso.com',
      accountType: 'member'
    },
    accountSettings: {
      usageLocation: 'SE',
      accountEnabled: true,
      credentialMode: 'password',
      forceChangePassword: true
    },
    licenses: [{ skuId: 'sku-1', skuPartNumber: 'SPE_E5' }],
    access: {
      securityGroups: [],
      m365Groups: [],
      teams: [],
      sharePointSites: [],
      applications: []
    },
    expirationReviewDays: null
  };
}

/** Polls a predicate on microtask ticks — used to synchronize with an in-flight mock call. */
async function waitUntil(predicate: () => boolean, maxTicks: number = 50): Promise<void> {
  for (let i = 0; i < maxTicks && !predicate(); i++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

interface IHarness {
  engine: WorkflowEngine;
  graph: MockGraph;
  data: MockData;
  credentials: ICredentialPresentation[];
}

function makeHarness(status: JobStatus, steps?: IJobStep[]): IHarness {
  const graph: MockGraph = new MockGraph();
  const data: MockData = new MockData(status, payload(), steps ?? []);
  const audit: AuditService = new AuditService(
    data as unknown as SharePointDataService,
    'operator@contoso.com'
  );
  const naming: NamingPolicyService = {
    checkUpnAvailability: async () => 'available'
  } as unknown as NamingPolicyService;
  const users: UserService = {
    isEmployeeIdTaken: async () => false
  } as unknown as UserService;
  const siteAccess: SiteAccessService = {
    grantAccess: async () => undefined
  } as unknown as SiteAccessService;
  const engine: WorkflowEngine = new WorkflowEngine(
    {
      graph: graph as unknown as GraphService,
      data: data as unknown as SharePointDataService,
      audit,
      naming,
      users,
      siteAccess
    },
    { stepBackoffBaseMs: 1 }
  );
  return { engine, graph, data, credentials: [] };
}

function happyPathHandlers(graph: MockGraph, options?: { failLicenses?: boolean }): void {
  graph.handlers = {
    // create-user pre-check by UPN, then POST
    'GET /users?$select=id&$filter=': () => ({ value: [] }),
    'POST /users/user-123/assignLicense': () => {
      if (options?.failLicenses) {
        throw new GraphServiceError('SKU not found', 400, 'Request_ResourceNotFound', 'req-lic-1');
      }
      return {};
    },
    'POST /users': () => ({ id: 'user-123' }),
    'GET /users/user-123?$select=id': () => ({ id: 'user-123' }),
    'GET /users/user-123?$select=usageLocation': () => ({ usageLocation: 'SE' }),
    'GET /users/user-123/manager': () => {
      throw new GraphServiceError('No manager', 404, 'Request_ResourceNotFound', 'req-mgr');
    },
    'PUT /users/user-123/manager/$ref': () => ({}),
    'GET /users/user-123/licenseDetails': () => ({ value: [] }),
    'PATCH /users/user-123': () => ({})
  };
}

// ---- tests --------------------------------------------------------------------

describe('WorkflowEngine', () => {
  it('refuses to start a job that is not yet approved (state machine)', async () => {
    const h: IHarness = makeHarness('PendingApproval');
    await expect(h.engine.runJob(1)).rejects.toThrow(/approval is required/);
    expect(h.data.status).toBe('PendingApproval');
    expect(h.graph.calls).toHaveLength(0);
  });

  it('refuses to start a Draft job', async () => {
    const h: IHarness = makeHarness('Draft');
    await expect(h.engine.runJob(1)).rejects.toThrow(/approval is required/);
  });

  it('runs an approved onboarding job to completion and hands over credentials once', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);
    const job: IProvisioningJob = await h.engine.runJob(1, {
      presentCredentials: async (c) => {
        h.credentials.push(c);
      }
    });

    expect(job.status).toBe('Completed');
    expect(h.data.status).toBe('Completed');
    expect(h.data.targetUserId).toBe('user-123');
    expect(h.graph.countCalls('POST', '/users')).toBe(2); // create + assignLicense
    expect(h.credentials).toHaveLength(1);
    expect(h.credentials[0].kind).toBe('password');
    expect(h.credentials[0].value.length).toBeGreaterThanOrEqual(16);
    // Secrets never persisted: not in steps, not in audit.
    expect(h.data.stepsJson).not.toContain(h.credentials[0].value);
    expect(JSON.stringify(h.data.auditEntries)).not.toContain(h.credentials[0].value);
    for (const step of h.data.steps()) {
      expect(step.status).toMatch(/completed|skipped/);
    }
  });

  it('marks a forced license failure Failed-Retryable with an audit entry carrying the response code', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph, { failLicenses: true });
    const job: IProvisioningJob = await h.engine.runJob(1, {
      presentCredentials: async () => undefined
    });

    // assign-licenses has continueOnFailure: false, so a blocking failure
    // here yields Failed, not PartiallyFailed (see the finalStatus comment
    // in WorkflowEngine._execute).
    expect(job.status).toBe('Failed');
    const licenseStep: IJobStep = h.data.steps().filter((s) => s.stepId === 'assign-licenses')[0];
    expect(licenseStep.status).toBe('failed');
    expect(licenseStep.lastError?.graphCode).toBe('Request_ResourceNotFound');
    const failureAudit: IAuditEntry[] = h.data.auditEntries.filter(
      (e) => e.result === 'Failure' && e.action === 'assign-licenses'
    );
    expect(failureAudit).toHaveLength(1);
    expect(failureAudit[0].responseCode).toBe(400);
    // The user was created exactly once before the failure.
    expect(h.graph.countCalls('POST', '/users')).toBe(2); // create + failed assignLicense
  });

  it('manual retry resumes without re-creating the user (idempotency)', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph, { failLicenses: true });
    await h.engine.runJob(1, { presentCredentials: async () => undefined });
    expect(h.data.status).toBe('Failed');

    // Operator fixes the underlying problem, then clicks Manual Retry.
    happyPathHandlers(h.graph, { failLicenses: false });
    const job: IProvisioningJob = await h.engine.retryStep(1, 'assign-licenses', {
      presentCredentials: async (c) => {
        h.credentials.push(c);
      }
    });

    expect(job.status).toBe('Completed');
    // POST /users (creation) happened exactly once across both runs.
    const userCreations: number = h.graph.calls.filter(
      (c) => c.method === 'POST' && c.path === '/users'
    ).length;
    expect(userCreations).toBe(1);
    // Password was re-issued via PATCH because the original never left memory.
    expect(h.graph.countCalls('PATCH', '/users/user-123')).toBe(1);
    expect(h.credentials).toHaveLength(1);
  });

  it('resumes from the first non-completed step after the tab was closed', async () => {
    // Simulate a job that died right after create-user completed.
    const h: IHarness = makeHarness('Running');
    happyPathHandlers(h.graph);
    const doneTemplate: Partial<IJobStep> = {
      status: 'completed',
      attempts: 1,
      maxAttempts: 3,
      lastError: null,
      startedUtc: '2026-07-04T08:00:00Z',
      completedUtc: '2026-07-04T08:00:05Z',
      skippable: false
    };
    const steps: IJobStep[] = [
      { ...doneTemplate, stepId: 'validate-input' } as IJobStep,
      { ...doneTemplate, stepId: 'create-user' } as IJobStep
    ];
    h.data.stepsJson = JSON.stringify(steps);
    h.data.targetUserId = 'user-123';

    const job: IProvisioningJob = await h.engine.runJob(1, {
      presentCredentials: async () => undefined
    });

    expect(job.status).toBe('Completed');
    // create-user never re-ran: no UPN pre-check, no POST /users creation.
    expect(h.graph.calls.filter((c) => c.method === 'POST' && c.path === '/users')).toHaveLength(0);
  });

  it('rejects a second concurrent run of the same job in this session', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);
    const first: Promise<IProvisioningJob> = h.engine.runJob(1, {
      presentCredentials: async () => undefined
    });
    await expect(h.engine.runJob(1)).rejects.toThrow(/already running/);
    await first;
  });

  it('skipStep refuses non-skippable steps', async () => {
    const h: IHarness = makeHarness('PartiallyFailed');
    const failed: IJobStep = {
      stepId: 'create-user',
      status: 'failed',
      attempts: 3,
      maxAttempts: 3,
      lastError: { graphCode: 'x', message: 'x', retryable: false },
      startedUtc: null,
      completedUtc: null,
      skippable: false
    };
    h.data.stepsJson = JSON.stringify([failed]);
    await expect(h.engine.skipStep(1, 'create-user')).rejects.toThrow(/not skippable/);
  });

  it('skip on a running step is not clobbered when its in-flight Graph call succeeds afterward', async () => {
    // Regression test: SPFx's Graph client gives no real request cancellation,
    // so a step's in-flight call can resolve successfully after skipStep()
    // already persisted 'skipped'. The engine must not overwrite that skip.
    // assign-manager is the only skippable onboarding step, so it stands in
    // for whatever in-flight step the operator clicks Skip on.
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);
    let resolveManagerLookup: (value: { id: string }) => void = () => undefined;
    let managerLookupStarted: boolean = false;
    h.graph.handlers['GET /users/user-123/manager'] = () => {
      managerLookupStarted = true;
      return new Promise<{ id: string }>((resolve) => {
        resolveManagerLookup = resolve;
      });
    };

    const runPromise: Promise<IProvisioningJob> = h.engine.runJob(1, {
      presentCredentials: async () => undefined
    });

    // Wait until the engine's manager lookup is actually in flight.
    await waitUntil(() => managerLookupStarted);
    expect(managerLookupStarted).toBe(true);

    // Operator clicks Skip on the currently-running, skippable step.
    await h.engine.skipStep(1, 'assign-manager');
    expect(h.data.steps().filter((s) => s.stepId === 'assign-manager')[0].status).toBe('skipped');

    // Now the in-flight lookup resolves successfully (a manager was assigned
    // out of band after all — PUT .../manager/$ref is stubbed to succeed).
    resolveManagerLookup({ id: 'someone-else' });
    const job: IProvisioningJob = await runPromise;

    // The explicit skip wins — it must not be clobbered back to 'completed'.
    const finalStep: IJobStep = job.steps.filter((s) => s.stepId === 'assign-manager')[0];
    expect(finalStep.status).toBe('skipped');
    const persistedStep: IJobStep = h.data.steps().filter((s) => s.stepId === 'assign-manager')[0];
    expect(persistedStep.status).toBe('skipped');
  });

  it('retryStep is rejected while the job is actively running', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);

    let releasePresentCredentials: () => void = () => undefined;
    const gate: Promise<void> = new Promise((resolve) => {
      releasePresentCredentials = resolve;
    });
    const runPromise: Promise<IProvisioningJob> = h.engine.runJob(1, {
      presentCredentials: async () => gate
    });

    // Wait until the job is registered as running in this session — retryStep
    // must reject purely because the job is running, regardless of which
    // step (or whether it's even in a retryable state) is named.
    await waitUntil(() => h.engine.isRunning(1));
    expect(h.engine.isRunning(1)).toBe(true);

    await expect(h.engine.retryStep(1, 'assign-manager')).rejects.toThrow(/currently running/);

    releasePresentCredentials();
    await runPromise;
  });
});
