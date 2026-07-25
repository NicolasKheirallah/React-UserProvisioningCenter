import { WorkflowEngine } from './WorkflowEngine';
import { GraphServiceError } from '../graph/GraphError';
import { AuditService } from '../audit/AuditService';
import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { NamingPolicyService } from '../namingPolicy/NamingPolicyService';
import type { UserService } from '../users/UserService';
import type { SiteAccessService } from '../sites/SiteAccessService';
import type { IAuthorizationService } from './IAuthorizationService';
import type { AppPermission, IAuditEntry, IJobStep, IJobPayload, IProvisioningJob, JobStatus, JobType } from '../../models';
import type { ICredentialPresentation } from './stepTypes';

export type Handler = (path: string, body?: unknown) => unknown;

export class MockGraph {
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

export class MockData {
  public status: JobStatus;
  public stepsJson: string;
  public targetUserId: string | null = null;
  public auditEntries: IAuditEntry[] = [];
  public lockOwner: string | undefined;
  private readonly _payload: IJobPayload;

  public constructor(status: JobStatus, payload: IJobPayload, steps: IJobStep[]) {
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
    batchId: '',
    targetUpn: 'anna.svensson@contoso.com',
    targetUserId: this.targetUserId,
    createdUtc: '2026-01-01T00:00:00Z',
    modifiedUtc: '2026-01-01T00:00:00Z',
    runningSince: null
  });

  public getJobStatus = async (_itemId: number): Promise<JobStatus> => this.status;

  public updateJobStatus = async (_itemId: number, status: JobStatus, _etag?: string): Promise<string> => {
    this.status = status;
    return '*';
  };

  public updateJobSteps = async (_itemId: number, steps: IJobStep[], _etag?: string): Promise<string> => {
    this.stepsJson = JSON.stringify(steps);
    return '*';
  };

  public setJobTargetUser = async (_itemId: number, userId: string, _etag?: string): Promise<string> => {
    this.targetUserId = userId;
    return '*';
  };

  public addAuditEntry = async (entry: IAuditEntry): Promise<void> => {
    this.auditEntries.push(entry);
  };

  public acquireJobLock = async (_itemId: number, instanceId: string): Promise<string> => {
    this.lockOwner = instanceId;
    return '*';
  };

  public releaseJobLock = async (_itemId: number, instanceId: string): Promise<void> => {
    if (this.lockOwner === instanceId) {
      this.lockOwner = undefined;
    }
  };

  public recordApproval = async (): Promise<{ satisfied: boolean; granted: number; required: number }> => {
    this.status = 'Approved';
    return { satisfied: true, granted: 1, required: 1 };
  };

  public steps(): IJobStep[] {
    return JSON.parse(this.stepsJson);
  }
}

export class AllowAllAuth implements IAuthorizationService {
  public denied: Set<AppPermission> = new Set();
  public async require(permission: AppPermission): Promise<void> {
    if (this.denied.has(permission)) {
      throw new Error(`Operation requires permission: ${permission}`);
    }
  }
  public async has(permission: AppPermission): Promise<boolean> {
    return !this.denied.has(permission);
  }
}

export function onboardingPayload(): IJobPayload {
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
  } as IJobPayload;
}

export async function waitUntil(predicate: () => boolean, maxTicks: number = 50): Promise<void> {
  for (let i = 0; i < maxTicks && !predicate(); i++) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

export interface IHarness {
  engine: WorkflowEngine;
  graph: MockGraph;
  data: MockData;
  auth: AllowAllAuth;
  credentials: ICredentialPresentation[];
}

export function makeHarness(status: JobStatus, steps?: IJobStep[], payload?: IJobPayload, jobType: JobType = 'Onboard'): IHarness {
  const graph: MockGraph = new MockGraph();
  const data: MockData = new MockData(status, payload ?? onboardingPayload(), steps ?? []);
  const audit: AuditService = new AuditService(data as unknown as SharePointDataService, 'operator@contoso.com');
  const naming: NamingPolicyService = { checkUpnAvailability: async () => 'available' } as unknown as NamingPolicyService;
  const users: UserService = { isEmployeeIdTaken: async () => false } as unknown as UserService;
  const siteAccess: SiteAccessService = { grantAccess: async () => undefined } as unknown as SiteAccessService;
  const auth: AllowAllAuth = new AllowAllAuth();
  const engine: WorkflowEngine = new WorkflowEngine(
    {
      graph: graph as unknown as GraphService,
      data: data as unknown as SharePointDataService,
      audit,
      naming,
      users,
      siteAccess,
      auth,
      operatorUpn: 'operator@contoso.com',
      operatorDisplayName: 'Operator'
    },
    { stepBackoffBaseMs: 1, cancelPollMinMs: 0 }
  );
  void jobType;
  return { engine, graph, data, auth, credentials: [] };
}
