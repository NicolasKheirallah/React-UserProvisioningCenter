import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/security';
import { PermissionKind } from '@pnp/sp/security';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  LIST_APPLICATION_CATALOG,
  LIST_APPROVAL_DELEGATIONS,
  LIST_AUDIT_LOG,
  LIST_DEPARTMENT_TEMPLATES,
  LIST_LICENSE_COST_TABLE,
  LIST_PROVISIONING_JOBS,
  LIST_ROLES,
  LIST_SETTINGS,
  LIST_SITE_CATALOG,
  LIST_TASKS,
  LIST_TEAMS_CATALOG
} from '../../constants/listNames';
import { UPC_LIST_DEFINITIONS } from '../provisioning/listSchemas';
import { escapeODataLiteral } from '../util/odata';
import { sharePointRetry } from '../util/sharePointRetry';
import { ConcurrencyError } from '../util/ConcurrencyError';
import { JobConflictError, isEtagConflict } from '../util/JobConflictError';
import { buildJobFilter } from '../util/jobFilter';
import type { IPagedResult } from '../util/pagedQuery';
import { fetchPaged } from '../util/pagedQuery';
import { assertTransition } from '../engine/jobStateMachine';
import type { TelemetryService } from '../telemetry/TelemetryService';
import type {
  ApplicationProvisioningType,
  AppPermission,
  AppRole,
  AuditResult,
  IAppSettings,
  IApplicationCatalogItem,
  IApprovalDelegation,
  IApprovalRecord,
  IAuditEntry,
  IDepartmentTemplate,
  IJobPayload,
  IJobQuery,
  IJobStep,
  IJobSummary,
  ILicenseCost,
  IProvisioningJob,
  IRoleDefinition,
  IRoleManagementItem,
  ISchemaGap,
  ISchemaValidationResult,
  IServiceDeskTask,
  ISiteCatalogItem,
  ITeamsCatalogItem,
  ITemplateListItem,
  JobStatus,
  JobType,
  TaskStatus,
  TaskType
} from '../../models';

interface IJobListItem {
  Id: number;
  Title: string;
  JobType: JobType;
  Status: JobStatus;
  PayloadJson: string;
  StepsJson: string;
  ScheduledFor: string | null;
  CorrelationId: string;
  TargetUpn: string | null;
  TargetUserId: string | null;
  ApprovalsJson: string | null;
  Created: string;
  Modified: string;
  RequestedBy?: { Title: string } | null;
  ApprovedBy?: { Title: string } | null;
}

const JOB_SELECT: string[] = [
  'Id',
  'Title',
  'JobType',
  'Status',
  'PayloadJson',
  'StepsJson',
  'ScheduledFor',
  'CorrelationId',
  'TargetUpn',
  'TargetUserId',
  'ApprovalsJson',
  'Created',
  'Modified',
  'RequestedBy/Title',
  'ApprovedBy/Title'
];

interface IJobSummaryListItem {
  Id: number;
  Title: string;
  JobType: JobType;
  Status: JobStatus;
  CorrelationId: string;
  TargetUpn: string | null;
  ScheduledFor: string | null;
  Created: string;
  Modified: string;
  RequestedBy?: { Title: string } | null;
  ApprovedBy?: { Title: string } | null;
}

const JOB_SUMMARY_SELECT: string[] = [
  'Id',
  'Title',
  'JobType',
  'Status',
  'CorrelationId',
  'TargetUpn',
  'ScheduledFor',
  'Created',
  'Modified',
  'RequestedBy/Title',
  'ApprovedBy/Title'
];

function parseJob(item: IJobListItem, telemetry?: TelemetryService): IProvisioningJob {
  let payload: IJobPayload;
  let steps: IJobStep[];
  let approvals: IApprovalRecord[];
  try {
    payload = JSON.parse(item.PayloadJson) as IJobPayload;
  } catch {
    payload = {} as IJobPayload;
    telemetry?.trackEvent(
      'data.parseFailure',
      { list: LIST_PROVISIONING_JOBS, field: 'PayloadJson', itemId: item.Id },
      'warning'
    );
  }
  try {
    steps = item.StepsJson ? (JSON.parse(item.StepsJson) as IJobStep[]) : [];
  } catch {
    steps = [];
    telemetry?.trackEvent(
      'data.parseFailure',
      { list: LIST_PROVISIONING_JOBS, field: 'StepsJson', itemId: item.Id },
      'warning'
    );
  }
  try {
    approvals = item.ApprovalsJson ? (JSON.parse(item.ApprovalsJson) as IApprovalRecord[]) : [];
  } catch {
    approvals = [];
    telemetry?.trackEvent(
      'data.parseFailure',
      { list: LIST_PROVISIONING_JOBS, field: 'ApprovalsJson', itemId: item.Id },
      'warning'
    );
  }
  return {
    itemId: item.Id,
    jobId: item.Title,
    jobType: item.JobType,
    status: item.Status,
    payload,
    steps,
    approvals,
    scheduledFor: item.ScheduledFor,
    requestedBy: item.RequestedBy?.Title ?? null,
    approvedBy: item.ApprovedBy?.Title ?? null,
    correlationId: item.CorrelationId,
    targetUpn: item.TargetUpn ?? '',
    targetUserId: item.TargetUserId,
    createdUtc: item.Created ?? null,
    modifiedUtc: item.Modified ?? null
  };
}

function parseJobSummary(item: IJobSummaryListItem): IJobSummary {
  return {
    itemId: item.Id,
    jobId: item.Title,
    jobType: item.JobType,
    status: item.Status,
    targetUpn: item.TargetUpn ?? '',
    scheduledFor: item.ScheduledFor,
    requestedBy: item.RequestedBy?.Title ?? null,
    approvedBy: item.ApprovedBy?.Title ?? null,
    correlationId: item.CorrelationId,
    createdUtc: item.Created ?? null,
    modifiedUtc: item.Modified ?? null
  };
}

export interface ICreateJobInput {
  jobId: string;
  jobType: JobType;
  payload: IJobPayload;
  steps: IJobStep[];
  scheduledFor: string | null;
  correlationId: string;
  targetUpn?: string;
  initialStatus?: 'PendingApproval' | 'Approved';
}

interface IDelegationListItem {
  Id: number;
  Title: string;
  DelegateUpn: string;
  StartUtc: string | null;
  EndUtc: string | null;
  Reason: string | null;
  IsActive: boolean;
}

function parseDelegation(item: IDelegationListItem): IApprovalDelegation {
  return {
    itemId: item.Id,
    delegatorUpn: item.Title,
    delegateUpn: item.DelegateUpn,
    startUtc: item.StartUtc,
    endUtc: item.EndUtc,
    reason: item.Reason ?? '',
    isActive: item.IsActive
  };
}

export class SharePointDataService {
  private readonly _sp: SPFI;
  private _telemetry: TelemetryService | undefined;
  private _currentUser: { Id: number; Title: string; LoginName: string } | undefined;

  public constructor(context: WebPartContext);
  public constructor(sp: SPFI);
  public constructor(contextOrSp: WebPartContext | SPFI) {
    this._sp =
      contextOrSp && typeof contextOrSp === 'object' && 'web' in contextOrSp
        ? (contextOrSp as SPFI)
        : spfi().using(SPFx(contextOrSp as WebPartContext));
  }

  public setTelemetry(telemetry: TelemetryService): void {
    this._telemetry = telemetry;
  }

  public static readonly DEFAULT_JOBS_TOP: number = 500;
  private static readonly DEFAULT_LOCK_STALE_MS: number = 10 * 60 * 1000;

  private _etagFromItem(item: unknown): string {
    return (item as Record<string, string>)['odata.etag'] ?? '*';
  }

  private async _me(): Promise<{ Id: number; Title: string; LoginName: string }> {
    if (!this._currentUser) {
      const fetched: { Id: number; Title: string; LoginName: string } = await sharePointRetry(
        () => this._sp.web.currentUser.select('Id', 'Title', 'LoginName')(),
        { circuitKey: 'sharepoint:currentUser' }
      );
      this._currentUser = fetched;
    }
    return this._currentUser;
  }

  private _jobs(): ReturnType<SPFI['web']['lists']['getByTitle']> {
    return this._sp.web.lists.getByTitle(LIST_PROVISIONING_JOBS);
  }

  private async _updateJob(
    itemId: number,
    fields: Record<string, unknown>,
    etag?: string
  ): Promise<string> {
    return sharePointRetry(
      async () => {
        const useEtag: string =
          etag ?? this._etagFromItem(await this._jobs().items.getById(itemId).select('odata.etag')());
        try {
          const updated = await this._jobs().items.getById(itemId).update(fields, useEtag);
          return this._etagFromItem(updated);
        } catch (err) {
          if (isEtagConflict(err)) {
            throw new JobConflictError(itemId);
          }
          throw err;
        }
      },
      { circuitKey: LIST_PROVISIONING_JOBS }
    );
  }

  public async getJobs(top: number = SharePointDataService.DEFAULT_JOBS_TOP): Promise<IProvisioningJob[]> {
    const page = await this.getJobsPaged(top);
    return page.items;
  }

  public async getJobsPaged(top: number = SharePointDataService.DEFAULT_JOBS_TOP): Promise<IPagedResult<IProvisioningJob>> {
    const query = this._jobs()
      .items.select(...JOB_SELECT, 'odata.etag')
      .expand('RequestedBy', 'ApprovedBy')
      .orderBy('Id', false)
      .top(top) as unknown as AsyncIterable<IJobListItem[]>;
    const page = await fetchPaged<IJobListItem>(query, top, (action) =>
      sharePointRetry(action, { circuitKey: LIST_PROVISIONING_JOBS })
    );
    return {
      items: page.items.map((item) => parseJob(item, this._telemetry)),
      truncated: page.truncated,
      next: page.next
        ? async () => {
            const nextPage = await page.next!();
            return {
              items: nextPage.items.map((item) => parseJob(item, this._telemetry)),
              truncated: nextPage.truncated
            };
          }
        : undefined
    };
  }

  public async getJobSummariesPaged(query?: IJobQuery): Promise<IPagedResult<IJobSummary>> {
    const top: number = query?.top ?? SharePointDataService.DEFAULT_JOBS_TOP;
    const filter: string = buildJobFilter(query);
    let request = this._jobs()
      .items.select(...JOB_SUMMARY_SELECT)
      .expand('RequestedBy', 'ApprovedBy')
      .orderBy('Id', false)
      .top(top);
    if (filter) {
      request = request.filter(filter);
    }
    const iterableRequest = request as unknown as AsyncIterable<IJobSummaryListItem[]>;
    const page = await fetchPaged<IJobSummaryListItem>(iterableRequest, top, (action) =>
      sharePointRetry(action, { circuitKey: LIST_PROVISIONING_JOBS })
    );
    return {
      items: page.items.map(parseJobSummary),
      truncated: page.truncated,
      next: page.next
        ? async () => {
            const nextPage = await page.next!();
            return { items: nextPage.items.map(parseJobSummary), truncated: nextPage.truncated };
          }
        : undefined
    };
  }

  public async getJobsChangeToken(): Promise<{ latestModifiedUtc: string; runningCount: number }> {
    return sharePointRetry(
      async () => {
        const newest: { Modified: string }[] = await this._jobs()
          .items.select('Modified')
          .orderBy('Modified', false)
          .top(1)();
        const running: { Id: number }[] = await this._jobs()
          .items.select('Id')
          .filter("Status eq 'Running'")
          .top(50)();
        return {
          latestModifiedUtc: newest[0]?.Modified ?? '',
          runningCount: running.length
        };
      },
      { circuitKey: LIST_PROVISIONING_JOBS }
    );
  }

  public async getJobStatus(itemId: number): Promise<JobStatus> {
    const item: { Status: JobStatus } = await sharePointRetry(
      () => this._jobs().items.getById(itemId).select('Status')(),
      { circuitKey: LIST_PROVISIONING_JOBS, maxAttempts: 2 }
    );
    return item.Status;
  }

  public async getJob(itemId: number): Promise<IProvisioningJob> {
    const item: IJobListItem = await sharePointRetry(
      () =>
        this._jobs()
          .items.getById(itemId)
          .select(...JOB_SELECT, 'odata.etag')
          .expand('RequestedBy', 'ApprovedBy')(),
      { circuitKey: LIST_PROVISIONING_JOBS }
    );
    return parseJob(item, this._telemetry);
  }

  public async createJob(input: ICreateJobInput): Promise<number> {
    const me = await this._me();
    return sharePointRetry(
      async () => {
        const result = await this._jobs().items.add({
          Title: input.jobId,
          JobType: input.jobType,
          Status: input.initialStatus ?? 'PendingApproval',
          PayloadJson: JSON.stringify(input.payload),
          StepsJson: JSON.stringify(input.steps),
          ScheduledFor: input.scheduledFor,
          CorrelationId: input.correlationId,
          TargetUpn: input.targetUpn ?? '',
          ApprovalsJson: '[]',
          RequestedById: me.Id
        });
        return (result as { Id: number }).Id;
      },
      { idempotent: false, circuitKey: LIST_PROVISIONING_JOBS }
    );
  }

  public async findJobIdByTitle(jobId: string): Promise<number | undefined> {
    const items: { Id: number }[] = await sharePointRetry(
      () => this._jobs().items.select('Id').filter(`Title eq '${escapeODataLiteral(jobId)}'`).top(1)(),
      { circuitKey: LIST_PROVISIONING_JOBS }
    );
    return items[0]?.Id;
  }

  public async updateJobStatus(
    itemId: number,
    status: JobStatus,
    etag?: string,
    options?: { skipTransitionCheck?: boolean }
  ): Promise<string> {
    if (!options?.skipTransitionCheck) {
      const current: { Status: JobStatus } = await sharePointRetry(
        () => this._jobs().items.getById(itemId).select('Status')(),
        { circuitKey: LIST_PROVISIONING_JOBS }
      );
      if (current.Status !== status) {
        assertTransition(current.Status, status);
      }
    }
    return this._updateJob(itemId, { Status: status }, etag);
  }

  public async approveJob(itemId: number, etag?: string): Promise<string> {
    const me = await this._me();
    const job: { Status: JobStatus } = await sharePointRetry(
      () => this._jobs().items.getById(itemId).select('Status')(),
      { circuitKey: LIST_PROVISIONING_JOBS }
    );
    assertTransition(job.Status, 'Approved');
    return this._updateJob(itemId, { Status: 'Approved', ApprovedById: me.Id }, etag);
  }

  public async recordApproval(
    itemId: number,
    approval: IApprovalRecord,
    requiredApprovals: number
  ): Promise<{ satisfied: boolean; granted: number; required: number }> {
    const job: IProvisioningJob = await this.getJob(itemId);
    const existing: IApprovalRecord[] = job.approvals ?? [];
    const alreadyByMe: boolean = existing.some(
      (a) => (a.actorUpn ?? '').toLowerCase() === approval.actorUpn.toLowerCase()
    );
    const approvals: IApprovalRecord[] = alreadyByMe ? existing : [...existing, approval];
    const distinct: number = new Set(
      approvals.map((a) => (a.actorUpn ?? '').toLowerCase()).filter((v) => v.length > 0)
    ).size;
    const required: number = Math.max(1, Math.floor(requiredApprovals) || 1);
    const satisfied: boolean = distinct >= required;

    const fields: Record<string, unknown> = { ApprovalsJson: JSON.stringify(approvals) };
    if (satisfied) {
      assertTransition(job.status, 'Approved');
      const me = await this._me();
      fields.Status = 'Approved';
      fields.ApprovedById = me.Id;
    }
    await this._updateJob(itemId, fields);
    return { satisfied, granted: distinct, required };
  }

  public async updateJobSteps(itemId: number, steps: IJobStep[], etag?: string): Promise<string> {
    return this._updateJob(itemId, { StepsJson: JSON.stringify(steps) }, etag);
  }

  public async setJobTargetUser(itemId: number, targetUserId: string, etag?: string): Promise<string> {
    return this._updateJob(itemId, { TargetUserId: targetUserId }, etag);
  }

  public async acquireJobLock(
    itemId: number,
    instanceId: string,
    maxLockAgeMs: number = SharePointDataService.DEFAULT_LOCK_STALE_MS
  ): Promise<string> {
    return sharePointRetry(
      async () => {
        const item: { RunningInstanceId?: string | null; RunningSince?: string | null; 'odata.etag'?: string } =
          await this._jobs().items.getById(itemId).select('Id', 'RunningInstanceId', 'RunningSince', 'odata.etag')();

        const currentLock: string | null = item.RunningInstanceId ?? null;
        const currentSince: string | null = item.RunningSince ?? null;
        const stale: boolean =
          !currentLock ||
          (currentSince !== undefined &&
            currentSince !== null &&
            Date.now() - new Date(currentSince).getTime() > maxLockAgeMs);

        if (currentLock && !stale) {
          throw new ConcurrencyError(`Job ${itemId} is already running in session ${currentLock}`);
        }

        const updated = await this._jobs()
          .items.getById(itemId)
          .update(
            { RunningInstanceId: instanceId, RunningSince: new Date().toISOString() },
            this._etagFromItem(item)
          );
        return this._etagFromItem(updated);
      },
      { circuitKey: LIST_PROVISIONING_JOBS }
    );
  }

  public async releaseJobLock(itemId: number, instanceId: string): Promise<void> {
    await sharePointRetry(
      async () => {
        const item: { RunningInstanceId?: string | null; 'odata.etag'?: string } = await this._jobs()
          .items.getById(itemId)
          .select('RunningInstanceId', 'odata.etag')();

        if (item.RunningInstanceId !== instanceId) {
          return;
        }

        await this._jobs()
          .items.getById(itemId)
          .update({ RunningInstanceId: null, RunningSince: null }, this._etagFromItem(item));
      },
      { circuitKey: LIST_PROVISIONING_JOBS }
    );
  }

  public async addAuditEntry(entry: IAuditEntry): Promise<void> {
    await sharePointRetry(
      () =>
        this._sp.web.lists.getByTitle(LIST_AUDIT_LOG).items.add({
          Title: entry.entryId,
          JobId: entry.jobId,
          Actor: entry.actor,
          Action: entry.action,
          TargetUser: entry.targetUser,
          GraphEndpoint: entry.graphEndpoint,
          RequestSummary: entry.requestSummary,
          ResponseCode: entry.responseCode,
          DurationMs: entry.durationMs,
          Result: entry.result,
          CorrelationId: entry.correlationId,
          TimestampUtc: entry.timestampUtc
        }),
      { idempotent: false, circuitKey: LIST_AUDIT_LOG }
    );
  }

  public async getAuditEntries(jobId: string, top: number = 100): Promise<IAuditEntry[]> {
    const literal: string = escapeODataLiteral(jobId);
    const items: {
      Title: string;
      JobId: string;
      Actor: string | null;
      Action: string;
      TargetUser: string | null;
      GraphEndpoint: string | null;
      RequestSummary: string | null;
      ResponseCode: number | null;
      DurationMs: number | null;
      Result: AuditResult;
      CorrelationId: string | null;
      TimestampUtc: string | null;
    }[] = await sharePointRetry(
      () =>
        this._sp.web.lists
          .getByTitle(LIST_AUDIT_LOG)
          .items.select(
            'Title',
            'JobId',
            'Actor',
            'Action',
            'TargetUser',
            'GraphEndpoint',
            'RequestSummary',
            'ResponseCode',
            'DurationMs',
            'Result',
            'CorrelationId',
            'TimestampUtc'
          )
          .filter(`JobId eq '${literal}'`)
          .orderBy('Id', true)
          .top(top)(),
      { circuitKey: LIST_AUDIT_LOG }
    );
    return items.map((i) => ({
      entryId: i.Title,
      jobId: i.JobId,
      actor: i.Actor ?? '',
      action: i.Action,
      targetUser: i.TargetUser ?? '',
      graphEndpoint: i.GraphEndpoint ?? '',
      requestSummary: i.RequestSummary ?? '',
      responseCode: i.ResponseCode ?? 0,
      durationMs: i.DurationMs ?? 0,
      result: i.Result,
      correlationId: i.CorrelationId ?? '',
      timestampUtc: i.TimestampUtc ?? ''
    }));
  }

  public async getRoleDefinitions(): Promise<IRoleDefinition[]> {
    const items: { Title: AppRole; MemberGroupId: string; PermissionsJson: string | null }[] = await sharePointRetry(
      () => this._sp.web.lists.getByTitle(LIST_ROLES).items.select('Title', 'MemberGroupId', 'PermissionsJson').top(50)(),
      { circuitKey: LIST_ROLES }
    );
    return items
      .filter((i) => !!i.MemberGroupId)
      .map((i) => {
        let permissions: AppPermission[] = [];
        try {
          permissions = i.PermissionsJson ? (JSON.parse(i.PermissionsJson) as AppPermission[]) : [];
        } catch {
          permissions = [];
          this._telemetry?.trackEvent('data.parseFailure', { list: 'UPC_Roles', field: 'PermissionsJson', role: i.Title }, 'warning');
        }
        return { role: i.Title, memberGroupId: i.MemberGroupId, permissions };
      });
  }

  public async getRoleDefinitionsForManagement(): Promise<IRoleManagementItem[]> {
    const items: { Id: number; Title: AppRole; MemberGroupId: string | null; PermissionsJson: string | null }[] =
      await sharePointRetry(
        () =>
          this._sp.web.lists
            .getByTitle(LIST_ROLES)
            .items.select('Id', 'Title', 'MemberGroupId', 'PermissionsJson')
            .top(50)(),
        { circuitKey: LIST_ROLES }
      );
    return items.map((i) => {
      let permissions: AppPermission[] = [];
      try {
        permissions = i.PermissionsJson ? (JSON.parse(i.PermissionsJson) as AppPermission[]) : [];
      } catch {
        permissions = [];
        this._telemetry?.trackEvent('data.parseFailure', { list: 'UPC_Roles', field: 'PermissionsJson', role: i.Title }, 'warning');
      }
      return { itemId: i.Id, role: i.Title, memberGroupId: i.MemberGroupId ?? '', permissions };
    });
  }

  public async updateRoleDefinition(itemId: number, memberGroupId: string, permissions: AppPermission[]): Promise<void> {
    await sharePointRetry(
      () =>
        this._sp.web.lists
          .getByTitle(LIST_ROLES)
          .items.getById(itemId)
          .update({ MemberGroupId: memberGroupId, PermissionsJson: JSON.stringify(permissions) }),
      { circuitKey: LIST_ROLES }
    );
  }

  public async getActiveDelegationsFor(delegateUpn: string): Promise<IApprovalDelegation[]> {
    const literal: string = escapeODataLiteral(delegateUpn);
    const items: IDelegationListItem[] = await sharePointRetry(
      () =>
        this._sp.web.lists
          .getByTitle(LIST_APPROVAL_DELEGATIONS)
          .items.select('Id', 'Title', 'DelegateUpn', 'StartUtc', 'EndUtc', 'Reason', 'IsActive')
          .filter(`IsActive eq 1 and DelegateUpn eq '${literal}'`)
          .top(50)(),
      { circuitKey: LIST_APPROVAL_DELEGATIONS }
    );
    return items.map(parseDelegation);
  }

  public async getAllDelegations(): Promise<IApprovalDelegation[]> {
    const items: IDelegationListItem[] = await sharePointRetry(
      () =>
        this._sp.web.lists
          .getByTitle(LIST_APPROVAL_DELEGATIONS)
          .items.select('Id', 'Title', 'DelegateUpn', 'StartUtc', 'EndUtc', 'Reason', 'IsActive')
          .top(200)(),
      { circuitKey: LIST_APPROVAL_DELEGATIONS }
    );
    return items.map(parseDelegation);
  }

  public async createDelegation(delegation: Omit<IApprovalDelegation, 'itemId'>): Promise<number> {
    const result = await sharePointRetry(
      () =>
        this._sp.web.lists.getByTitle(LIST_APPROVAL_DELEGATIONS).items.add({
          Title: delegation.delegatorUpn,
          DelegateUpn: delegation.delegateUpn,
          StartUtc: delegation.startUtc,
          EndUtc: delegation.endUtc,
          Reason: delegation.reason,
          IsActive: delegation.isActive
        }),
      { idempotent: false, circuitKey: LIST_APPROVAL_DELEGATIONS }
    );
    return (result as { Id: number }).Id;
  }

  public async setDelegationActive(itemId: number, isActive: boolean): Promise<void> {
    await sharePointRetry(
      () =>
        this._sp.web.lists
          .getByTitle(LIST_APPROVAL_DELEGATIONS)
          .items.getById(itemId)
          .update({ IsActive: isActive }),
      { circuitKey: LIST_APPROVAL_DELEGATIONS }
    );
  }

  public async getLicenseCosts(): Promise<ILicenseCost[]> {
    const items: { Title: string; MonthlyCost: number; Currency: string }[] = await sharePointRetry(
      () => this._sp.web.lists.getByTitle(LIST_LICENSE_COST_TABLE).items.select('Title', 'MonthlyCost', 'Currency').top(500)(),
      { circuitKey: LIST_LICENSE_COST_TABLE }
    );
    return items.map((i) => ({ skuPartNumber: i.Title, monthlyCost: i.MonthlyCost, currency: i.Currency }));
  }

  public async getTeamsCatalog(): Promise<ITeamsCatalogItem[]> {
    const items: { Id: number; Title: string; TeamId: string; Category: string | null; DefaultRole: string | null }[] =
      await sharePointRetry(
        () =>
          this._sp.web.lists
            .getByTitle(LIST_TEAMS_CATALOG)
            .items.select('Id', 'Title', 'TeamId', 'Category', 'DefaultRole')
            .top(500)(),
        { circuitKey: LIST_TEAMS_CATALOG }
      );
    return items.map((i) => ({
      itemId: i.Id,
      title: i.Title,
      teamId: i.TeamId,
      category: i.Category ?? '',
      defaultRole: i.DefaultRole === 'owner' ? 'owner' : 'member'
    }));
  }

  public async getSiteCatalog(): Promise<ISiteCatalogItem[]> {
    const items: {
      Id: number;
      Title: string;
      SiteUrl: string;
      BusinessOwner?: { Title: string } | null;
      Category: string | null;
    }[] = await sharePointRetry(
      () =>
        this._sp.web.lists
          .getByTitle(LIST_SITE_CATALOG)
          .items.select('Id', 'Title', 'SiteUrl', 'BusinessOwner/Title', 'Category')
          .expand('BusinessOwner')
          .top(500)(),
      { circuitKey: LIST_SITE_CATALOG }
    );
    return items.map((i) => ({
      itemId: i.Id,
      title: i.Title,
      siteUrl: i.SiteUrl,
      businessOwner: i.BusinessOwner?.Title ?? null,
      category: i.Category ?? ''
    }));
  }

  public async getApplicationCatalog(): Promise<IApplicationCatalogItem[]> {
    const items: {
      Id: number;
      Title: string;
      Owner?: { Title: string } | null;
      ProvisioningType: string | null;
      TargetGroupId: string | null;
      ApprovalRequired: boolean | null;
      Instructions: string | null;
      IsActive: boolean | null;
    }[] = await sharePointRetry(
      () =>
        this._sp.web.lists
          .getByTitle(LIST_APPLICATION_CATALOG)
          .items.select(
            'Id',
            'Title',
            'Owner/Title',
            'ProvisioningType',
            'TargetGroupId',
            'ApprovalRequired',
            'Instructions',
            'IsActive'
          )
          .expand('Owner')
          .filter('IsActive eq 1')
          .top(500)(),
      { circuitKey: LIST_APPLICATION_CATALOG }
    );
    return items.map((i) => ({
      itemId: i.Id,
      title: i.Title,
      owner: i.Owner?.Title ?? null,
      provisioningType: (i.ProvisioningType === 'GroupBased' ? 'GroupBased' : 'Manual') as ApplicationProvisioningType,
      targetGroupId: i.TargetGroupId,
      approvalRequired: !!i.ApprovalRequired,
      instructions: i.Instructions ?? '',
      isActive: i.IsActive !== false
    }));
  }

  public async getActiveTemplates(): Promise<ITemplateListItem[]> {
    const items: { Id: number; Title: string; TemplateJson: string; IsActive: boolean; Version: number }[] =
      await sharePointRetry(
        () =>
          this._sp.web.lists
            .getByTitle(LIST_DEPARTMENT_TEMPLATES)
            .items.select('Id', 'Title', 'TemplateJson', 'IsActive', 'Version')
            .filter('IsActive eq 1')
            .top(200)(),
        { circuitKey: LIST_DEPARTMENT_TEMPLATES }
      );
    const templates: ITemplateListItem[] = [];
    for (const i of items) {
      try {
        templates.push({
          itemId: i.Id,
          title: i.Title,
          template: JSON.parse(i.TemplateJson),
          isActive: i.IsActive,
          version: i.Version
        });
      } catch (parseErr) {
        void parseErr;
      }
    }
    return templates;
  }

  public async getAllTemplates(): Promise<ITemplateListItem[]> {
    const items: { Id: number; Title: string; TemplateJson: string; IsActive: boolean; Version: number }[] =
      await sharePointRetry(
        () =>
          this._sp.web.lists
            .getByTitle(LIST_DEPARTMENT_TEMPLATES)
            .items.select('Id', 'Title', 'TemplateJson', 'IsActive', 'Version')
            .top(200)(),
        { circuitKey: LIST_DEPARTMENT_TEMPLATES }
      );
    return items.map((i) => {
      let template: IDepartmentTemplate;
      try {
        template = JSON.parse(i.TemplateJson) as IDepartmentTemplate;
      } catch {
        this._telemetry?.trackEvent('data.parseFailure', { list: 'UPC_DepartmentTemplates', field: 'TemplateJson', itemId: i.Id }, 'warning');
        template = {
          department: '',
          licenses: [],
          securityGroups: [],
          m365Groups: [],
          teams: [],
          sharePointSites: [],
          applications: [],
          approverGroupId: null,
          expirationPolicyDays: null,
          usageLocationDefault: ''
        };
      }
      template.approverGroupId = template.approverGroupId ?? null;
      return { itemId: i.Id, title: i.Title, template, isActive: i.IsActive, version: i.Version };
    });
  }

  public async createTemplate(title: string, template: IDepartmentTemplate): Promise<number> {
    const result = await sharePointRetry(
      () =>
        this._sp.web.lists.getByTitle(LIST_DEPARTMENT_TEMPLATES).items.add({
          Title: title,
          TemplateJson: JSON.stringify(template),
          IsActive: true,
          Version: 1
        }),
      { idempotent: false, circuitKey: LIST_DEPARTMENT_TEMPLATES }
    );
    return (result as { Id: number }).Id;
  }

  public async updateTemplate(itemId: number, title: string, template: IDepartmentTemplate, currentVersion: number): Promise<void> {
    await sharePointRetry(
      () =>
        this._sp.web.lists
          .getByTitle(LIST_DEPARTMENT_TEMPLATES)
          .items.getById(itemId)
          .update({ Title: title, TemplateJson: JSON.stringify(template), Version: currentVersion + 1 }),
      { circuitKey: LIST_DEPARTMENT_TEMPLATES }
    );
  }

  public async setTemplateActive(itemId: number, isActive: boolean): Promise<void> {
    await sharePointRetry(
      () => this._sp.web.lists.getByTitle(LIST_DEPARTMENT_TEMPLATES).items.getById(itemId).update({ IsActive: isActive }),
      { circuitKey: LIST_DEPARTMENT_TEMPLATES }
    );
  }

  public async getAppSettings(): Promise<Partial<IAppSettings>> {
    const items: { Id: number; SettingsJson: string | null }[] = await sharePointRetry(
      () => this._sp.web.lists.getByTitle(LIST_SETTINGS).items.select('Id', 'SettingsJson').filter("Title eq 'app'").top(1)(),
      { circuitKey: LIST_SETTINGS }
    );
    if (items.length === 0) {
      return {};
    }
    try {
      return JSON.parse(items[0].SettingsJson ?? '{}') as Partial<IAppSettings>;
    } catch {
      this._telemetry?.trackEvent('data.parseFailure', { list: 'UPC_Settings', field: 'SettingsJson', itemId: items[0].Id }, 'warning');
      return {};
    }
  }

  public async saveAppSettings(settings: IAppSettings): Promise<void> {
    await sharePointRetry(
      async () => {
        const list = this._sp.web.lists.getByTitle(LIST_SETTINGS);
        const items: { Id: number }[] = await list.items.select('Id').filter("Title eq 'app'").top(1)();
        const payload = { Title: 'app', SettingsJson: JSON.stringify(settings) };
        if (items.length > 0) {
          await list.items.getById(items[0].Id).update(payload);
        } else {
          await list.items.add(payload);
        }
      },
      { circuitKey: LIST_SETTINGS }
    );
  }

  public async getTasks(top: number = 500): Promise<IServiceDeskTask[]> {
    const page = await this.getTasksPaged(top);
    return page.items;
  }

  public async getTasksPaged(top: number = 500): Promise<IPagedResult<IServiceDeskTask>> {
    const query = this._sp.web.lists
      .getByTitle(LIST_TASKS)
      .items.select('Id', 'Title', 'JobId', 'TaskType', 'Instructions', 'Status', 'CompletedBy', 'CompletedUtc', 'AssignedTo/Title')
      .expand('AssignedTo')
      .orderBy('Id', false)
      .top(top) as unknown as AsyncIterable<
      {
        Id: number;
        Title: string;
        JobId: string | null;
        TaskType: TaskType;
        Instructions: string | null;
        Status: TaskStatus;
        CompletedBy: string | null;
        CompletedUtc: string | null;
        AssignedTo?: { Title: string } | null;
      }[]
    >;
    const toTask = (i: {
      Id: number;
      Title: string;
      JobId: string | null;
      TaskType: TaskType;
      Instructions: string | null;
      Status: TaskStatus;
      CompletedBy: string | null;
      CompletedUtc: string | null;
      AssignedTo?: { Title: string } | null;
    }): IServiceDeskTask => ({
      itemId: i.Id,
      title: i.Title,
      jobId: i.JobId ?? '',
      taskType: i.TaskType,
      instructions: i.Instructions ?? '',
      assignedTo: i.AssignedTo?.Title ?? null,
      status: i.Status,
      completedBy: i.CompletedBy ?? null,
      completedUtc: i.CompletedUtc ?? null
    });
    const page = await fetchPaged(query, top, (action) => sharePointRetry(action, { circuitKey: LIST_TASKS }));
    return {
      items: page.items.map(toTask),
      truncated: page.truncated,
      next: page.next
        ? async () => {
            const nextPage = await page.next!();
            return { items: nextPage.items.map(toTask), truncated: nextPage.truncated };
          }
        : undefined
    };
  }

  public async completeTask(itemId: number): Promise<void> {
    const me = await this._me();
    await sharePointRetry(
      () =>
        this._sp.web.lists
          .getByTitle(LIST_TASKS)
          .items.getById(itemId)
          .update({ Status: 'Done', CompletedBy: me.Title, CompletedUtc: new Date().toISOString() }),
      { circuitKey: LIST_TASKS }
    );
  }

  public async createTask(jobId: string, taskType: TaskType, title: string, instructions: string): Promise<void> {
    await sharePointRetry(
      () =>
        this._sp.web.lists.getByTitle(LIST_TASKS).items.add({
          Title: title,
          JobId: jobId,
          TaskType: taskType,
          Instructions: instructions,
          Status: 'Open'
        }),
      { idempotent: false, circuitKey: LIST_TASKS }
    );
  }

  public async probeWriteAccess(): Promise<boolean> {
    return sharePointRetry(
      () => this._sp.web.lists.getByTitle(LIST_TASKS).currentUserHasPermissions(PermissionKind.AddListItems),
      { circuitKey: LIST_TASKS }
    );
  }

  public async validateSchema(): Promise<ISchemaValidationResult> {
    const results: (ISchemaGap | undefined)[] = await Promise.all(
      UPC_LIST_DEFINITIONS.map(async (definition) => {
        try {
          const fields: { InternalName: string }[] = await sharePointRetry(
            () => this._sp.web.lists.getByTitle(definition.title).fields.select('InternalName')(),
            { circuitKey: definition.title, maxAttempts: 2 }
          );
          const present: Set<string> = new Set(fields.map((f) => f.InternalName));
          const missingFields: string[] = definition.fields
            .map((f) => f.name)
            .filter((name) => !present.has(name));
          if (missingFields.length === 0) {
            return undefined;
          }
          return { list: definition.title, missingList: false, missingFields, error: '' };
        } catch (err) {
          const status: number =
            typeof err === 'object' && err !== null && 'status' in err
              ? Number((err as { status: unknown }).status)
              : 0;
          return {
            list: definition.title,
            missingList: status === 404,
            missingFields: [],
            error: err instanceof Error ? err.message : String(err)
          };
        }
      })
    );
    return {
      gaps: results.filter((g): g is ISchemaGap => g !== undefined),
      checkedLists: UPC_LIST_DEFINITIONS.length
    };
  }
}
