import { GraphServiceError } from '../graph/GraphError';
import type { IAuditEntry, IJobStep, IProvisioningJob } from '../../models';
import { makeHarness, waitUntil, type IHarness, type MockGraph } from './testHarness';

function happyPathHandlers(graph: MockGraph, options?: { failLicenses?: boolean }): void {
  graph.handlers = {
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

describe('WorkflowEngine', () => {
  it('refuses to start a job that is not yet approved (state machine)', async () => {
    const h: IHarness = makeHarness('PendingApproval');
    await expect(h.engine.runJob(1)).rejects.toThrow(/approval is required/);
    expect(h.data.status).toBe('PendingApproval');
    expect(h.graph.calls).toHaveLength(0);
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
    expect(h.graph.countCalls('POST', '/users')).toBe(2);
    expect(h.credentials).toHaveLength(1);
    expect(h.credentials[0].kind).toBe('password');
    expect(h.credentials[0].value.length).toBeGreaterThanOrEqual(16);
    expect(h.data.stepsJson).not.toContain(h.credentials[0].value);
    expect(JSON.stringify(h.data.auditEntries)).not.toContain(h.credentials[0].value);
    for (const step of h.data.steps()) {
      expect(step.status).toMatch(/completed|skipped/);
    }
  });

  it('releases the job lock only after the run has fully completed, not before', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);
    const lockDuringRun: boolean[] = [];
    const originalUpdateJobSteps = h.data.updateJobSteps;
    h.data.updateJobSteps = async (itemId, steps, etag) => {
      lockDuringRun.push(h.data.lockOwner !== undefined);
      return originalUpdateJobSteps(itemId, steps, etag);
    };

    await h.engine.runJob(1, { presentCredentials: async () => undefined });

    expect(lockDuringRun.length).toBeGreaterThan(0);
    expect(lockDuringRun.every(Boolean)).toBe(true);
    expect(h.data.lockOwner).toBeUndefined();
  });

  it('marks a forced license failure Failed with an audit entry carrying the response code', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph, { failLicenses: true });
    const job: IProvisioningJob = await h.engine.runJob(1, { presentCredentials: async () => undefined });

    expect(job.status).toBe('Failed');
    const licenseStep: IJobStep = h.data.steps().filter((s) => s.stepId === 'assign-licenses')[0];
    expect(licenseStep.status).toBe('failed');
    expect(licenseStep.lastError?.graphCode).toBe('Request_ResourceNotFound');
    const failureAudit: IAuditEntry[] = h.data.auditEntries.filter(
      (e) => e.result === 'Failure' && e.action === 'assign-licenses'
    );
    expect(failureAudit).toHaveLength(1);
    expect(failureAudit[0].responseCode).toBe(400);
    expect(h.graph.countCalls('POST', '/users')).toBe(2);
  });

  it('manual retry resumes without re-creating the user (idempotency)', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph, { failLicenses: true });
    await h.engine.runJob(1, { presentCredentials: async () => undefined });
    expect(h.data.status).toBe('Failed');

    happyPathHandlers(h.graph, { failLicenses: false });
    const job: IProvisioningJob = await h.engine.retryStep(1, 'assign-licenses', {
      presentCredentials: async (c) => {
        h.credentials.push(c);
      }
    });

    expect(job.status).toBe('Completed');
    const userCreations: number = h.graph.calls.filter((c) => c.method === 'POST' && c.path === '/users').length;
    expect(userCreations).toBe(1);
    expect(h.graph.countCalls('PATCH', '/users/user-123')).toBe(1);
    expect(h.credentials).toHaveLength(1);
  });

  it('resumes from the first non-completed step after the tab was closed', async () => {
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

    const job: IProvisioningJob = await h.engine.runJob(1, { presentCredentials: async () => undefined });

    expect(job.status).toBe('Completed');
    expect(h.graph.calls.filter((c) => c.method === 'POST' && c.path === '/users')).toHaveLength(0);
  });

  it('rejects a second concurrent run of the same job in this session', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);
    const first: Promise<IProvisioningJob> = h.engine.runJob(1, { presentCredentials: async () => undefined });
    await expect(h.engine.runJob(1)).rejects.toThrow(/already running/);
    await first;
  });

  it('stops mid-pipeline when another session cancels the job', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);

    const originalGetStatus = h.data.getJobStatus;
    let probes: number = 0;
    h.data.getJobStatus = async (itemId: number) => {
      probes++;
      if (probes > 1) {
        return 'Cancelled';
      }
      return originalGetStatus(itemId);
    };

    const job: IProvisioningJob = await h.engine.runJob(1, { presentCredentials: async () => undefined });

    expect(job.status).toBe('Cancelled');
    expect(job.steps.some((s) => s.status === 'pending')).toBe(true);
  });

  it('runs to completion when no cancellation is requested', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);

    const job: IProvisioningJob = await h.engine.runJob(1, { presentCredentials: async () => undefined });

    expect(job.status).toBe('Completed');
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

    const runPromise: Promise<IProvisioningJob> = h.engine.runJob(1, { presentCredentials: async () => undefined });

    await waitUntil(() => managerLookupStarted);
    expect(managerLookupStarted).toBe(true);

    await h.engine.skipStep(1, 'assign-manager');
    expect(h.data.steps().filter((s) => s.stepId === 'assign-manager')[0].status).toBe('skipped');

    resolveManagerLookup({ id: 'someone-else' });
    const job: IProvisioningJob = await runPromise;

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
    const runPromise: Promise<IProvisioningJob> = h.engine.runJob(1, { presentCredentials: async () => gate });

    await waitUntil(() => h.engine.isRunning(1));
    expect(h.engine.isRunning(1)).toBe(true);

    await expect(h.engine.retryStep(1, 'assign-manager')).rejects.toThrow(/currently running/);

    releasePresentCredentials();
    await runPromise;
  });

  it('createJob requires the createJobs permission', async () => {
    const h: IHarness = makeHarness('PendingApproval');
    h.auth.denied.add('createJobs');
    await expect(
      h.engine.createJob({ jobType: 'Onboard', payload: {} as never, steps: [], scheduledFor: null })
    ).rejects.toThrow(/createJobs/);
  });

  it('approveJob records a distinct approver and satisfies a 1-of-1 quorum', async () => {
    const h: IHarness = makeHarness('PendingApproval');
    const outcome = await h.engine.approveJob(1);
    expect(outcome.satisfied).toBe(true);
    expect(outcome.granted).toBe(1);
    expect(h.data.status).toBe('Approved');
  });
});
