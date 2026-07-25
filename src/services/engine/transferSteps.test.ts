import { WorkflowEngine } from './WorkflowEngine';
import { GraphServiceError } from '../graph/GraphError';
import { AuditService } from '../audit/AuditService';
import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { NamingPolicyService } from '../namingPolicy/NamingPolicyService';
import type { UserService } from '../users/UserService';
import type { SiteAccessService } from '../sites/SiteAccessService';
import type { IAuthorizationService } from './IAuthorizationService';
import type { IAuditEntry, IJobStep, IProvisioningJob, ITransferPayload, JobStatus } from '../../models';

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
  private readonly _payload: ITransferPayload;

  public constructor(status: JobStatus, payload: ITransferPayload, steps: IJobStep[]) {
    this.status = status;
    this._payload = payload;
    this.stepsJson = JSON.stringify(steps);
  }

  public getJob = async (itemId: number): Promise<IProvisioningJob> => ({
    itemId,
    jobId: 'job-guid-1',
    jobType: 'Transfer',
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

function payload(): ITransferPayload {
  return {
    schemaVersion: 1,
    kind: 'transfer',
    target: {
      userId: 'user-123',
      displayName: 'Anna Svensson',
      userPrincipalName: 'anna.svensson@contoso.com'
    },
    changes: {
      jobTitle: 'Senior Controller',
      department: 'Group Finance',
      managerId: 'mgr-1',
      managerDisplayName: 'Mgr One',
      addLicenses: [{ skuId: 'sku-2', skuPartNumber: 'SPE_F1' }],
      removeLicenseSkuIds: ['sku-1']
    }
  };
}

interface IHarness {
  engine: WorkflowEngine;
  graph: MockGraph;
  data: MockData;
}

function makeHarness(status: JobStatus, txPayload: ITransferPayload = payload()): IHarness {
  const graph = new MockGraph();
  const data = new MockData(status, txPayload, []);
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
    'GET /users/user-123?$select=id': () => ({ id: 'user-123' }),
    'GET /users/user-123?$select=jobTitle,department,officeLocation': () => ({
      jobTitle: 'Controller',
      department: 'Finance',
      officeLocation: 'Stockholm'
    }),
    'PATCH /users/user-123': () => ({}),
    'GET /users/user-123/manager': () => ({ id: 'someone-else' }),
    'PUT /users/user-123/manager/$ref': () => ({}),
    'GET /users/user-123/licenseDetails': () => ({ value: [{ skuId: 'sku-1' }] }),
    'POST /users/user-123/assignLicense': () => ({})
  };
}

describe('transfer: end-to-end pipeline', () => {
  it('applies only the diffed employment fields, updates the manager, and reconciles licenses', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);

    const job: IProvisioningJob = await h.engine.runJob(1);

    expect(job.status).toBe('Completed');
    expect(h.data.targetUserId).toBe('user-123');

    const employmentPatch = h.graph.calls.filter(
      (c) => c.method === 'PATCH' && c.path === '/users/user-123'
    )[0];
    expect(employmentPatch.body).toEqual({ jobTitle: 'Senior Controller', department: 'Group Finance' });

    expect(h.graph.calls.some((c) => c.method === 'PUT' && c.path === '/users/user-123/manager/$ref')).toBe(
      true
    );

    const licenseCall = h.graph.calls.filter(
      (c) => c.method === 'POST' && c.path === '/users/user-123/assignLicense'
    )[0];
    expect(licenseCall.body).toEqual({
      addLicenses: [{ skuId: 'sku-2', disabledPlans: [] }],
      removeLicenses: ['sku-1']
    });
    for (const step of h.data.steps()) {
      expect(step.status).toMatch(/completed|skipped/);
    }
  });

  it('skips the manager PUT when the current manager already matches (idempotent)', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);
    h.graph.handlers['GET /users/user-123/manager'] = () => ({ id: 'mgr-1' });

    await h.engine.runJob(1);

    expect(h.graph.calls.some((c) => c.method === 'PUT' && c.path === '/users/user-123/manager/$ref')).toBe(
      false
    );
  });

  it('fails validate-transfer (blocking) when the target user is not found', async () => {
    const h: IHarness = makeHarness('Approved');
    h.graph.handlers = {
      'GET /users/user-123?$select=id': () => {
        throw new GraphServiceError('Not found', 404, 'Request_ResourceNotFound', 'req-tgt');
      }
    };

    const job: IProvisioningJob = await h.engine.runJob(1);

    expect(job.status).toBe('Failed');
    const step: IJobStep = h.data.steps().filter((s) => s.stepId === 'validate-transfer')[0];
    expect(step.status).toBe('failed');
    expect(step.lastError?.graphCode).toBe('UPC_TargetNotFound');
  });

  it('fails validate-transfer (blocking) when the payload specifies no changes at all', async () => {
    const h: IHarness = makeHarness('Approved', {
      schemaVersion: 1,
      kind: 'transfer',
      target: {
        userId: 'user-123',
        displayName: 'Anna Svensson',
        userPrincipalName: 'anna.svensson@contoso.com'
      },
      changes: { addLicenses: [], removeLicenseSkuIds: [] }
    });
    h.graph.handlers = { 'GET /users/user-123?$select=id': () => ({ id: 'user-123' }) };

    const job: IProvisioningJob = await h.engine.runJob(1);

    expect(job.status).toBe('Failed');
    const step: IJobStep = h.data.steps().filter((s) => s.stepId === 'validate-transfer')[0];
    expect(step.lastError?.graphCode).toBe('UPC_InvalidPayload');
  });
});

describe('transfer: send-notifications', () => {
  it('emails the current manager a summary once the transfer completes', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);
    h.graph.handlers['GET /users/user-123/manager'] = () => ({
      id: 'mgr-1',
      mail: 'manager@contoso.com'
    });
    h.graph.handlers['POST /me/sendMail'] = () => ({});

    const job: IProvisioningJob = await h.engine.runJob(1);

    expect(job.status).toBe('Completed');
    const mailCall = h.graph.calls.filter((c) => c.method === 'POST' && c.path === '/me/sendMail')[0];
    expect(mailCall).toBeDefined();
    const body = mailCall.body as { message: { toRecipients: { emailAddress: { address: string } }[] } };
    expect(body.message.toRecipients[0].emailAddress.address).toBe('manager@contoso.com');
  });

  it('skips silently when the user has no manager on file', async () => {
    const h: IHarness = makeHarness('Approved');
    happyPathHandlers(h.graph);
    h.graph.handlers['GET /users/user-123/manager'] = () => {
      throw new GraphServiceError('Not found', 404, 'Request_ResourceNotFound', 'req-mgr');
    };

    const job: IProvisioningJob = await h.engine.runJob(1);

    expect(job.status).toBe('Completed');
    expect(h.graph.countCalls('POST', '/me/sendMail')).toBe(0);
  });
});
