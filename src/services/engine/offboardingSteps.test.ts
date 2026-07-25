import { WorkflowEngine } from './WorkflowEngine';
import { GraphServiceError } from '../graph/GraphError';
import { AuditService } from '../audit/AuditService';
import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { NamingPolicyService } from '../namingPolicy/NamingPolicyService';
import type { UserService } from '../users/UserService';
import type { SiteAccessService } from '../sites/SiteAccessService';
import type { IAuthorizationService } from './IAuthorizationService';
import type { IAuditEntry, IJobStep, IOffboardingPayload, IProvisioningJob, JobStatus, TaskType } from '../../models';

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
  public batch = async (
    requests: { id: string; method: string; url: string; body?: unknown }[]
  ): Promise<Map<string, { id: string; status: number; body: unknown }>> => {
    const map = new Map<string, { id: string; status: number; body: unknown }>();
    for (const req of requests) {
      try {
        const body = await this._dispatch(req.method, req.url, req.body);
        map.set(req.id, { id: req.id, status: 200, body });
      } catch (err) {
        const status = err instanceof GraphServiceError ? err.statusCode : 500;
        map.set(req.id, { id: req.id, status, body: {} });
      }
    }
    return map;
  };

  public countCalls(method: string, pathPrefix: string): number {
    return this.calls.filter((c) => c.method === method && c.path.startsWith(pathPrefix)).length;
  }
}

class MockData {
  public status: JobStatus;
  public stepsJson: string;
  public targetUserId: string | null = null;
  public auditEntries: IAuditEntry[] = [];
  public tasks: { jobId: string; taskType: TaskType; title: string; instructions: string }[] = [];
  private readonly _payload: IOffboardingPayload;

  public constructor(status: JobStatus, payload: IOffboardingPayload, steps: IJobStep[]) {
    this.status = status;
    this._payload = payload;
    this.stepsJson = JSON.stringify(steps);
  }

  public getJob = async (itemId: number): Promise<IProvisioningJob> => ({
    itemId,
    jobId: 'job-guid-1',
    jobType: 'Offboard',
    status: this.status,
    payload: JSON.parse(JSON.stringify(this._payload)),
    steps: JSON.parse(this.stepsJson),
    approvals: [],
    scheduledFor: null,
    requestedBy: 'HR Person',
    approvedBy: null,
    correlationId: 'corr-1',
    targetUpn: '',
    targetUserId: this.targetUserId,
    createdUtc: '2026-01-01T00:00:00Z',
    modifiedUtc: '2026-01-01T00:00:00Z'
  });

  public updateJobStatus = async (_itemId: number, status: JobStatus): Promise<string> => {
    this.status = status;
    return '*';
  };

  public updateJobSteps = async (_itemId: number, steps: IJobStep[]): Promise<string> => {
    this.stepsJson = JSON.stringify(steps);
    return '*';
  };

  public setJobTargetUser = async (_itemId: number, userId: string): Promise<string> => {
    this.targetUserId = userId;
    return '*';
  };

  public addAuditEntry = async (entry: IAuditEntry): Promise<void> => {
    this.auditEntries.push(entry);
  };

  public acquireJobLock = async (): Promise<string> => '*';
  public releaseJobLock = async (): Promise<void> => undefined;

  public createTask = async (
    jobId: string,
    taskType: TaskType,
    title: string,
    instructions: string
  ): Promise<void> => {
    this.tasks.push({ jobId, taskType, title, instructions });
  };

  public steps(): IJobStep[] {
    return JSON.parse(this.stepsJson);
  }
}

function payload(): IOffboardingPayload {
  return {
    schemaVersion: 1,
    kind: 'offboard',
    identity: {
      userPrincipalName: 'anna.svensson@contoso.com',
      mailNickname: 'anna.svensson',
      domain: 'contoso.com',
      accountType: 'member'
    },
    target: {
      userId: 'user-123',
      displayName: 'Anna Svensson',
      userPrincipalName: 'anna.svensson@contoso.com'
    },
    options: {
      removeLicenses: false,
      removeFromGroups: false,
      mailboxAction: 'none'
    }
  };
}

interface IHarness {
  engine: WorkflowEngine;
  graph: MockGraph;
  data: MockData;
}

function makeHarness(status: JobStatus): IHarness {
  const graph = new MockGraph();
  const data = new MockData(status, payload(), []);
  const audit = new AuditService(data as unknown as SharePointDataService, 'operator@contoso.com');
  const naming = { checkUpnAvailability: async () => 'available' } as unknown as NamingPolicyService;
  const users = { isEmployeeIdTaken: async () => false } as unknown as UserService;
  const siteAccess = { grantAccess: async () => undefined } as unknown as SiteAccessService;
  const engine = new WorkflowEngine(
    {
      graph: graph as unknown as GraphService,
      data: data as unknown as SharePointDataService,
      audit,
      naming,
      users,
      siteAccess,
      auth: { require: async () => undefined } as unknown as IAuthorizationService,
      operatorUpn: 'operator@contoso.com'
    },
    { stepBackoffBaseMs: 1 }
  );
  return { engine, graph, data };
}

function baseHandlers(graph: MockGraph, onPremisesSyncEnabled: boolean | null): void {
  graph.handlers = {
    'GET /users/user-123?$select=id': () => ({ id: 'user-123' }),
    'GET /users/user-123?$select=accountEnabled,onPremisesSyncEnabled': () => ({
      accountEnabled: true,
      onPremisesSyncEnabled
    }),
    'POST /users/user-123/revokeSignInSessions': () => ({}),
    'PATCH /users/user-123': () => ({}),
    'GET /users/user-123/memberOf/microsoft.graph.group': () => ({ value: [] }),
    'GET /users/user-123/licenseDetails': () => ({ value: [] })
  };
}

describe('offboarding: block-sign-in', () => {
  it('disables sign-in directly through Graph for a cloud-only user', async () => {
    const h: IHarness = makeHarness('Approved');
    baseHandlers(h.graph, null);

    const job: IProvisioningJob = await h.engine.runJob(1);

    expect(job.status).toBe('Completed');
    expect(h.graph.calls.some((c) => c.method === 'PATCH' && c.path === '/users/user-123')).toBe(true);
    expect(h.graph.countCalls('POST', '/users/user-123/revokeSignInSessions')).toBe(1);
    expect(h.data.tasks).toHaveLength(0);
  });

  it('does not PATCH accountEnabled for a directory-synced user — hands off to on-prem AD instead', async () => {
    const h: IHarness = makeHarness('Approved');
    baseHandlers(h.graph, true);

    const job: IProvisioningJob = await h.engine.runJob(1);

    expect(job.status).toBe('Completed');
    // The synced-attribute write must never be attempted.
    expect(h.graph.calls.some((c) => c.method === 'PATCH' && c.path === '/users/user-123')).toBe(false);
    // Session revocation still happens — it isn't attribute-sync-dependent.
    expect(h.graph.countCalls('POST', '/users/user-123/revokeSignInSessions')).toBe(1);
    // A task hands the on-prem action off to whoever manages AD.
    expect(h.data.tasks).toHaveLength(1);
    expect(h.data.tasks[0].taskType).toBe('OnPremAdAccount');
    expect(h.data.tasks[0].instructions).toContain('on-prem');
  });
});

describe('offboarding: send-notifications', () => {
  it('skips silently when no hand-over contact was recorded', async () => {
    const h: IHarness = makeHarness('Approved');
    baseHandlers(h.graph, null);

    const job: IProvisioningJob = await h.engine.runJob(1);

    expect(job.status).toBe('Completed');
    expect(h.graph.countCalls('POST', '/me/sendMail')).toBe(0);
  });

  it('emails the OneDrive hand-over contact once the job completes', async () => {
    const graph = new MockGraph();
    const withOneDrive: IOffboardingPayload = {
      ...payload(),
      options: { ...payload().options, oneDriveAccessUpn: 'manager@contoso.com' }
    };
    const data = new MockData('Approved', withOneDrive, []);
    const audit = new AuditService(data as unknown as SharePointDataService, 'operator@contoso.com');
    const engine = new WorkflowEngine(
      {
        graph: graph as unknown as GraphService,
        data: data as unknown as SharePointDataService,
        audit,
        naming: { checkUpnAvailability: async () => 'available' } as unknown as NamingPolicyService,
        users: { isEmployeeIdTaken: async () => false } as unknown as UserService,
        siteAccess: { grantAccess: async () => undefined } as unknown as SiteAccessService,
        auth: { require: async () => undefined } as unknown as IAuthorizationService,
        operatorUpn: 'operator@contoso.com'
      },
      { stepBackoffBaseMs: 1 }
    );
    baseHandlers(graph, null);
    graph.handlers['POST /me/sendMail'] = () => ({});

    const job: IProvisioningJob = await engine.runJob(1);

    expect(job.status).toBe('Completed');
    const mailCall = graph.calls.filter((c) => c.method === 'POST' && c.path === '/me/sendMail')[0];
    expect(mailCall).toBeDefined();
    const body = mailCall.body as { message: { toRecipients: { emailAddress: { address: string } }[] } };
    expect(body.message.toRecipients[0].emailAddress.address).toBe('manager@contoso.com');
  });
});
