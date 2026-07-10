import { WorkflowEngine } from './WorkflowEngine';
import { GraphServiceError } from '../graph/GraphError';
import { AuditService } from '../audit/AuditService';
import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { NamingPolicyService } from '../namingPolicy/NamingPolicyService';
import type { UserService } from '../users/UserService';
import type { SiteAccessService } from '../sites/SiteAccessService';
import type { IAuditEntry, IOnboardingPayload, IProvisioningJob, JobStatus, JobType } from '../../models';
import type { ICredentialPresentation } from './stepTypes';

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
  public auditEntries: IAuditEntry[] = [];

  public constructor(
    private readonly _status: JobStatus,
    private readonly _jobType: JobType,
    private readonly _payload: IOnboardingPayload,
    private readonly _targetUserId: string | null
  ) {}

  public getJob = async (itemId: number): Promise<IProvisioningJob> => ({
    itemId,
    jobId: 'job-guid-1',
    jobType: this._jobType,
    status: this._status,
    payload: JSON.parse(JSON.stringify(this._payload)),
    steps: [],
    scheduledFor: null,
    requestedBy: 'HR Person',
    approvedBy: null,
    correlationId: 'corr-1',
    targetUserId: this._targetUserId,
    createdUtc: '2026-01-01T00:00:00Z'
  });

  public addAuditEntry = async (entry: IAuditEntry): Promise<void> => {
    this.auditEntries.push(entry);
  };
}

function payload(overrides?: Partial<IOnboardingPayload>): IOnboardingPayload {
  return {
    schemaVersion: 1,
    kind: 'onboard',
    personal: { firstName: 'Anna', lastName: 'Svensson', displayName: 'Anna Svensson', employeeId: 'E12345' },
    employment: { jobTitle: 'Controller', department: 'Finance', employeeType: 'Employee', hireDate: '2026-08-01' },
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
    licenses: [],
    access: { securityGroups: [], m365Groups: [], teams: [], sharePointSites: [], applications: [] },
    expirationReviewDays: null,
    ...overrides
  };
}

function makeEngine(
  graph: MockGraph,
  data: MockData
): WorkflowEngine {
  const audit = new AuditService(data as unknown as SharePointDataService, 'operator@contoso.com');
  const naming = { checkUpnAvailability: async () => 'available' } as unknown as NamingPolicyService;
  const users = { isEmployeeIdTaken: async () => false } as unknown as UserService;
  const siteAccess = { grantAccess: async () => undefined } as unknown as SiteAccessService;
  return new WorkflowEngine(
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
}

describe('regenerateCredentials', () => {
  it('issues a new password via PATCH and presents it once', async () => {
    const graph = new MockGraph();
    graph.handlers = { 'PATCH /users/user-123': () => ({}) };
    const data = new MockData('Completed', 'Onboard', payload(), 'user-123');
    const engine = makeEngine(graph, data);
    const presented: ICredentialPresentation[] = [];

    await engine.regenerateCredentials(1, {
      presentCredentials: async (c) => {
        presented.push(c);
      }
    });

    expect(graph.countCalls('PATCH', '/users/user-123')).toBe(1);
    const patchBody = graph.calls[0].body as { passwordProfile: { password: string } };
    expect(patchBody.passwordProfile.password.length).toBeGreaterThanOrEqual(16);
    expect(presented).toHaveLength(1);
    expect(presented[0].kind).toBe('password');
    expect(presented[0].value).toBe(patchBody.passwordProfile.password);
    expect(data.auditEntries).toHaveLength(1);
    expect(data.auditEntries[0].action).toBe('regenerate-credentials');
    expect(data.auditEntries[0].result).toBe('Success');
  });

  it('drops any stale TAP before issuing a new one in TAP mode', async () => {
    const graph = new MockGraph();
    graph.handlers = {
      'GET /users/user-123/authentication/temporaryAccessPassMethods': () => ({
        value: [{ id: 'tap-old' }]
      }),
      'DELETE /users/user-123/authentication/temporaryAccessPassMethods/tap-old': () => ({}),
      'POST /users/user-123/authentication/temporaryAccessPassMethods': () => ({
        temporaryAccessPass: 'a-fresh-tap-value'
      })
    };
    const data = new MockData(
      'Completed',
      'Clone',
      payload({ accountSettings: { usageLocation: 'SE', accountEnabled: true, credentialMode: 'tap', forceChangePassword: false } }),
      'user-123'
    );
    const engine = makeEngine(graph, data);
    const presented: ICredentialPresentation[] = [];

    await engine.regenerateCredentials(1, {
      presentCredentials: async (c) => {
        presented.push(c);
      }
    });

    expect(
      graph.calls.some(
        (c) => c.method === 'DELETE' && c.path === '/users/user-123/authentication/temporaryAccessPassMethods/tap-old'
      )
    ).toBe(true);
    expect(presented[0].kind).toBe('tap');
    expect(presented[0].value).toBe('a-fresh-tap-value');
  });

  it('refuses when the job is not yet completed', async () => {
    const graph = new MockGraph();
    const data = new MockData('Running', 'Onboard', payload(), 'user-123');
    const engine = makeEngine(graph, data);

    await expect(engine.regenerateCredentials(1)).rejects.toThrow(/completed/);
    expect(graph.calls).toHaveLength(0);
  });

  it('refuses for a guest invite (no password or TAP exists)', async () => {
    const graph = new MockGraph();
    const data = new MockData(
      'Completed',
      'Onboard',
      payload({ identity: { userPrincipalName: 'guest@fabrikam.com', mailNickname: '', domain: '', accountType: 'guest' } }),
      'guest-123'
    );
    const engine = makeEngine(graph, data);

    await expect(engine.regenerateCredentials(1)).rejects.toThrow(/[Gg]uest/);
    expect(graph.calls).toHaveLength(0);
  });

  it('refuses for job types other than Onboard/Clone', async () => {
    const graph = new MockGraph();
    const data = new MockData('Completed', 'Offboard', payload(), 'user-123');
    const engine = makeEngine(graph, data);

    await expect(engine.regenerateCredentials(1)).rejects.toThrow(/onboarding and clone/);
  });

  it('records a Failure audit entry and rethrows when the Graph write fails', async () => {
    const graph = new MockGraph();
    graph.handlers = {
      'PATCH /users/user-123': () => {
        throw new GraphServiceError('Forbidden', 403, 'Authorization_RequestDenied', 'req-1');
      }
    };
    const data = new MockData('Completed', 'Onboard', payload(), 'user-123');
    const engine = makeEngine(graph, data);

    await expect(engine.regenerateCredentials(1)).rejects.toThrow(/Forbidden/);
    expect(data.auditEntries).toHaveLength(1);
    expect(data.auditEntries[0].result).toBe('Failure');
    expect(data.auditEntries[0].responseCode).toBe(403);
  });
});
