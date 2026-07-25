import { WorkflowEngine } from './WorkflowEngine';
import { GraphServiceError } from '../graph/GraphError';
import { AuditService } from '../audit/AuditService';
import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { NamingPolicyService } from '../namingPolicy/NamingPolicyService';
import type { UserService } from '../users/UserService';
import type { SiteAccessService } from '../sites/SiteAccessService';
import type { IAuthorizationService } from './IAuthorizationService';
import type { IAuditEntry, IJobStep, IOnboardingPayload, IProvisioningJob, JobStatus } from '../../models';

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
    approvals: [],
    scheduledFor: null,
    requestedBy: 'HR Person',
    approvedBy: null,
    correlationId: 'corr-1',
    targetUpn: '',
    targetUserId: this.targetUserId,
    createdUtc: '2026-01-01T00:00:00Z',
    modifiedUtc: '2026-01-01T00:00:00Z',
    runningSince: null
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

  public steps(): IJobStep[] {
    return JSON.parse(this.stepsJson);
  }
}

function guestPayload(): IOnboardingPayload {
  return {
    schemaVersion: 1,
    kind: 'onboard',
    personal: {
      firstName: 'Guest',
      lastName: 'Partner',
      displayName: 'Guest Partner',
      employeeId: 'EXT-001'
    },
    employment: {
      jobTitle: 'External Consultant',
      department: 'Finance',
      employeeType: 'Contractor',
      hireDate: '2026-08-01'
    },
    identity: {

      userPrincipalName: 'guest.partner@fabrikam.com',
      mailNickname: '',
      domain: '',
      accountType: 'guest'
    },
    accountSettings: {
      usageLocation: 'SE',
      accountEnabled: true,
      credentialMode: 'password',
      forceChangePassword: false
    },
    licenses: [],
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

interface IHarness {
  engine: WorkflowEngine;
  graph: MockGraph;
  data: MockData;
}

function makeHarness(status: JobStatus): IHarness {
  const graph = new MockGraph();
  const data = new MockData(status, guestPayload(), []);
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

function happyPathHandlers(graph: MockGraph): void {
  graph.handlers = {
    'GET /users?$select=id&$filter=': () => ({ value: [] }),
    'POST /invitations': () => ({ invitedUser: { id: 'guest-123' } })
  };
}

describe('onboarding: guest invite', () => {
  it('invites via /invitations instead of creating a cloud account, and skips license/credential steps', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);
    let presentedCredentials: boolean = false;

    const job: IProvisioningJob = await h.engine.runJob(1, {
      presentCredentials: async () => {
        presentedCredentials = true;
      }
    });

    expect(job.status).toBe('Completed');
    expect(h.data.targetUserId).toBe('guest-123');
    expect(h.graph.calls.some((c) => c.method === 'POST' && c.path === '/invitations')).toBe(true);

    expect(h.graph.calls.some((c) => c.method === 'POST' && c.path === '/users')).toBe(false);
    expect(h.graph.calls.some((c) => c.method === 'POST' && c.path === '/users/guest-123/assignLicense')).toBe(
      false
    );
    expect(presentedCredentials).toBe(false);
    for (const step of h.data.steps()) {
      expect(step.status).toMatch(/completed|skipped/);
    }
  });

  it('is idempotent when the invitation already created the directory object (response lost)', async () => {
    const h: IHarness = makeHarness('Approved');
    let invitationCalls: number = 0;
    h.graph.handlers = {
      'GET /users?$select=id&$filter=': () => ({ value: [{ id: 'guest-123' }] }),
      'POST /invitations': () => {
        invitationCalls++;
        return { invitedUser: { id: 'guest-999' } };
      }
    };

    const job: IProvisioningJob = await h.engine.runJob(1, {
      presentCredentials: async () => undefined
    });

    expect(job.status).toBe('Completed');

    expect(h.data.targetUserId).toBe('guest-123');
    expect(invitationCalls).toBe(0);
  });
});
