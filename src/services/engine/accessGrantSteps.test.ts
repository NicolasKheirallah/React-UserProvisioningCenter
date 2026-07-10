import { WorkflowEngine } from './WorkflowEngine';
import { GraphServiceError } from '../graph/GraphError';
import { AuditService } from '../audit/AuditService';
import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { NamingPolicyService } from '../namingPolicy/NamingPolicyService';
import type { UserService } from '../users/UserService';
import type { SiteAccessService } from '../sites/SiteAccessService';
import type {
  IApplicationCatalogItem,
  IAuditEntry,
  IJobStep,
  IOnboardingPayload,
  IProvisioningJob,
  JobStatus,
  TaskType
} from '../../models';

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
  public tasks: { jobId: string; taskType: TaskType; title: string; instructions: string }[] = [];
  private readonly _payload: IOnboardingPayload;
  private readonly _applicationCatalog: IApplicationCatalogItem[];

  public constructor(
    status: JobStatus,
    payload: IOnboardingPayload,
    steps: IJobStep[],
    applicationCatalog: IApplicationCatalogItem[] = []
  ) {
    this.status = status;
    this._payload = payload;
    this.stepsJson = JSON.stringify(steps);
    this._applicationCatalog = applicationCatalog;
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

  public createTask = async (
    jobId: string,
    taskType: TaskType,
    title: string,
    instructions: string
  ): Promise<void> => {
    this.tasks.push({ jobId, taskType, title, instructions });
  };

  public getApplicationCatalog = async (): Promise<IApplicationCatalogItem[]> => this._applicationCatalog;

  public steps(): IJobStep[] {
    return JSON.parse(this.stepsJson);
  }
}

function basePayload(): IOnboardingPayload {
  return {
    schemaVersion: 1,
    kind: 'onboard',
    personal: {
      firstName: 'Elin',
      lastName: 'Berg',
      displayName: 'Elin Berg',
      employeeId: 'E99001'
    },
    employment: {
      jobTitle: 'Engineer',
      department: 'IT',
      employeeType: 'Employee',
      hireDate: '2026-08-01'
    },
    identity: {
      userPrincipalName: 'elin.berg@contoso.com',
      mailNickname: 'elin.berg',
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
    access: {
      securityGroups: ['grp-sec-1'],
      m365Groups: ['grp-365-1'],
      teams: [{ teamId: 'team-1', role: 'member' }],
      sharePointSites: [
        { siteUrl: 'https://contoso.sharepoint.com/sites/it-ok', role: 'member' },
        { siteUrl: 'https://contoso.sharepoint.com/sites/it-fail', role: 'visitor' }
      ],
      applications: ['1', '2']
    },
    expirationReviewDays: null
  };
}

const APPLICATION_CATALOG: IApplicationCatalogItem[] = [
  {
    itemId: 1,
    title: 'Salesforce',
    owner: 'it-ops@contoso.com',
    provisioningType: 'GroupBased',
    targetGroupId: 'grp-app-target',
    approvalRequired: false,
    instructions: '',
    isActive: true
  },
  {
    itemId: 2,
    title: 'Legacy ERP',
    owner: 'it-ops@contoso.com',
    provisioningType: 'Manual',
    targetGroupId: null,
    approvalRequired: true,
    instructions: 'Provision through the vendor portal.',
    isActive: true
  }
];

interface IHarness {
  engine: WorkflowEngine;
  graph: MockGraph;
  data: MockData;
  siteAccessCalls: { siteUrl: string; upn: string; role: string }[];
}

function makeHarness(status: JobStatus, siteAccessFails: (siteUrl: string) => boolean = () => false): IHarness {
  const graph = new MockGraph();
  const data = new MockData(status, basePayload(), [], APPLICATION_CATALOG);
  const audit = new AuditService(data as unknown as SharePointDataService, 'operator@contoso.com');
  const naming = { checkUpnAvailability: async () => 'available' } as unknown as NamingPolicyService;
  const users = { isEmployeeIdTaken: async () => false } as unknown as UserService;
  const siteAccessCalls: { siteUrl: string; upn: string; role: string }[] = [];
  const siteAccess = {
    grantAccess: async (siteUrl: string, upn: string, role: string) => {
      siteAccessCalls.push({ siteUrl, upn, role });
      if (siteAccessFails(siteUrl)) {
        throw new Error('access denied');
      }
    }
  } as unknown as SiteAccessService;
  const engine = new WorkflowEngine(
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
  return { engine, graph, data, siteAccessCalls };
}

function happyPathHandlers(graph: MockGraph, options?: { failOneGroup?: boolean; failOneTeam?: boolean }): void {
  graph.handlers = {
    'GET /users?$select=id&$filter=': () => ({ value: [] }),
    'POST /users': () => ({ id: 'user-123' }),
    'GET /users/user-123?$select=id': () => ({ id: 'user-123' }),
    'GET /users/user-123?$select=usageLocation': () => ({ usageLocation: 'SE' }),
    'GET /users/user-123/manager': () => {
      throw new GraphServiceError('No manager', 404, 'Request_ResourceNotFound', 'req-mgr');
    },
    'GET /users/user-123/licenseDetails': () => ({ value: [] }),
    'PATCH /users/user-123': () => ({}),
    'POST /users/user-123/checkMemberGroups': () => ({ value: [] }),
    'POST /groups/grp-sec-1/members/$ref': () => ({}),
    'POST /groups/grp-365-1/members/$ref': () => {
      if (options?.failOneGroup) {
        throw new GraphServiceError('Forbidden', 403, 'Authorization_RequestDenied', 'req-grp');
      }
      return {};
    },
    'POST /groups/grp-app-target/members/$ref': () => ({}),
    'GET /teams/team-1/members': () => ({ value: [] }),
    'POST /teams/team-1/members': () => {
      if (options?.failOneTeam) {
        throw new GraphServiceError('Forbidden', 403, 'Authorization_RequestDenied', 'req-team');
      }
      return {};
    }
  };
}

describe('onboarding: access-grant steps', () => {
  it('grants security/M365 groups, Teams, SharePoint sites, and group-based applications; hands off failures/manual apps as tasks', async () => {
    const h: IHarness = makeHarness('Approved', (siteUrl) => siteUrl.endsWith('it-fail'));
    happyPathHandlers(h.graph);

    const job: IProvisioningJob = await h.engine.runJob(1, {
      presentCredentials: async () => undefined
    });

    // Each access-grant step catches its own per-item failures and hands them
    // off as a task rather than throwing — the step itself still completes.
    expect(job.status).toBe('Completed');
    expect(h.graph.calls.some((c) => c.method === 'POST' && c.path === '/groups/grp-sec-1/members/$ref')).toBe(
      true
    );
    expect(h.graph.calls.some((c) => c.method === 'POST' && c.path === '/groups/grp-365-1/members/$ref')).toBe(
      true
    );
    expect(h.graph.calls.some((c) => c.method === 'POST' && c.path === '/teams/team-1/members')).toBe(true);
    // One SharePoint site succeeded, one failed — both went through SiteAccessService.
    expect(h.siteAccessCalls.map((c) => c.siteUrl)).toEqual([
      'https://contoso.sharepoint.com/sites/it-ok',
      'https://contoso.sharepoint.com/sites/it-fail'
    ]);
    expect(h.data.tasks.some((t) => t.taskType === 'SharePointAccess')).toBe(true);
    // Group-based app joined the target group; the manual app produced a task.
    expect(
      h.graph.calls.some((c) => c.method === 'POST' && c.path === '/groups/grp-app-target/members/$ref')
    ).toBe(true);
    const appTask = h.data.tasks.filter((t) => t.taskType === 'ApplicationAssignment')[0];
    expect(appTask).toBeDefined();
    expect(appTask.instructions).toContain('Legacy ERP');
  });

  it('sends a GroupAssignment task when a group add is denied by Graph', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph, { failOneGroup: true });

    await h.engine.runJob(1, { presentCredentials: async () => undefined });

    const groupTask = h.data.tasks.filter((t) => t.taskType === 'GroupAssignment')[0];
    expect(groupTask).toBeDefined();
    expect(groupTask.instructions).toContain('grp-365-1');
  });

  it('sends a TeamAssignment task when a Teams add is denied by Graph', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph, { failOneTeam: true });

    await h.engine.runJob(1, { presentCredentials: async () => undefined });

    const teamTask = h.data.tasks.filter((t) => t.taskType === 'TeamAssignment')[0];
    expect(teamTask).toBeDefined();
    expect(teamTask.instructions).toContain('team-1');
  });
});
