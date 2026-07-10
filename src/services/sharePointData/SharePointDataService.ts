import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  LIST_APPLICATION_CATALOG,
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
import { escapeODataLiteral } from '../util/odata';
import type { TelemetryService } from '../telemetry/TelemetryService';
import type {
  ApplicationProvisioningType,
  AppPermission,
  AppRole,
  AuditResult,
  IAppSettings,
  IApplicationCatalogItem,
  IAuditEntry,
  IDepartmentTemplate,
  IJobPayload,
  IJobStep,
  ILicenseCost,
  IProvisioningJob,
  IRoleDefinition,
  IRoleManagementItem,
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
  TargetUserId: string | null;
  Created: string;
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
  'TargetUserId',
  'Created',
  'RequestedBy/Title',
  'ApprovedBy/Title'
];

function parseJob(item: IJobListItem, telemetry?: TelemetryService): IProvisioningJob {
  let payload: IJobPayload;
  let steps: IJobStep[];
  try {
    payload = JSON.parse(item.PayloadJson) as IJobPayload;
  } catch {
    payload = {} as IJobPayload;
    telemetry?.trackEvent('data.parseFailure', { list: 'UPC_Jobs', field: 'PayloadJson', itemId: item.Id }, 'warning');
  }
  try {
    steps = item.StepsJson ? (JSON.parse(item.StepsJson) as IJobStep[]) : [];
  } catch {
    steps = [];
    telemetry?.trackEvent('data.parseFailure', { list: 'UPC_Jobs', field: 'StepsJson', itemId: item.Id }, 'warning');
  }
  return {
    itemId: item.Id,
    jobId: item.Title,
    jobType: item.JobType,
    status: item.Status,
    payload,
    steps,
    scheduledFor: item.ScheduledFor,
    requestedBy: item.RequestedBy?.Title ?? null,
    approvedBy: item.ApprovedBy?.Title ?? null,
    correlationId: item.CorrelationId,
    targetUserId: item.TargetUserId,
    createdUtc: item.Created ?? null
  };
}

export interface ICreateJobInput {
  jobId: string;
  jobType: JobType;
  payload: IJobPayload;
  steps: IJobStep[];
  scheduledFor: string | null;
  correlationId: string;
  /**
   * PendingApproval (default) when the approval gate is on; Approved when the
   * web part is configured to skip approval — the job is immediately runnable.
   */
  initialStatus?: 'PendingApproval' | 'Approved';
}

/** Page result that flags when the underlying query was capped by `top`. */
export interface IPagedResult<T> {
  items: T[];
  /** True when the server returned exactly `top` rows — more may exist. */
  truncated: boolean;
}

/**
 * All SharePoint list access (PnPJS v4). Configuration, job state and audit
 * live in UPC_* lists on the host site (spec Section 4).
 */
export class SharePointDataService {
  private readonly _sp: SPFI;
  private _telemetry: TelemetryService | undefined;

  public constructor(context: WebPartContext);
  public constructor(sp: SPFI);
  public constructor(contextOrSp: WebPartContext | SPFI) {
    // A pre-built SPFI shares one PnPJS root across services; a web part
    // context creates a standalone root (legacy/property-pane callers).
    this._sp =
      contextOrSp && typeof contextOrSp === 'object' && 'web' in contextOrSp
        ? (contextOrSp as SPFI)
        : spfi().using(SPFx(contextOrSp as WebPartContext));
  }

  /** Set post-construction — TelemetryService is created after this service in the composition root. */
  public setTelemetry(telemetry: TelemetryService): void {
    this._telemetry = telemetry;
  }

  // ---- Jobs ----------------------------------------------------------------

  /** Default page size; large enough for most tenants, small enough to keep payload reasonable. */
  public static readonly DEFAULT_JOBS_TOP: number = 500;

  public async getJobs(top: number = SharePointDataService.DEFAULT_JOBS_TOP): Promise<IProvisioningJob[]> {
    const items: IJobListItem[] = await this._sp.web.lists
      .getByTitle(LIST_PROVISIONING_JOBS)
      .items.select(...JOB_SELECT)
      .expand('RequestedBy', 'ApprovedBy')
      .orderBy('Modified', false)
      .top(top)();
    return items.map((item) => parseJob(item, this._telemetry));
  }

  /** Paged variant that reports truncation so callers can surface a warning. */
  public async getJobsPaged(top: number = SharePointDataService.DEFAULT_JOBS_TOP): Promise<IPagedResult<IProvisioningJob>> {
    // Request top+1 to detect whether more rows exist without a false positive
    // when the list has exactly `top` items.
    const items: IJobListItem[] = await this._sp.web.lists
      .getByTitle(LIST_PROVISIONING_JOBS)
      .items.select(...JOB_SELECT)
      .expand('RequestedBy', 'ApprovedBy')
      .orderBy('Modified', false)
      .top(top + 1)();
    const truncated: boolean = items.length > top;
    const pageItems: IJobListItem[] = truncated ? items.slice(0, top) : items;
    return { items: pageItems.map((item) => parseJob(item, this._telemetry)), truncated };
  }

  public async getJob(itemId: number): Promise<IProvisioningJob> {
    const item: IJobListItem = await this._sp.web.lists
      .getByTitle(LIST_PROVISIONING_JOBS)
      .items.getById(itemId)
      .select(...JOB_SELECT)
      .expand('RequestedBy', 'ApprovedBy')();
    return parseJob(item, this._telemetry);
  }

  public async createJob(input: ICreateJobInput): Promise<number> {
    const me: { Id: number } = await this._sp.web.currentUser.select('Id')();
    const result = await this._sp.web.lists.getByTitle(LIST_PROVISIONING_JOBS).items.add({
      Title: input.jobId,
      JobType: input.jobType,
      Status: input.initialStatus ?? 'PendingApproval',
      PayloadJson: JSON.stringify(input.payload),
      StepsJson: JSON.stringify(input.steps),
      ScheduledFor: input.scheduledFor,
      CorrelationId: input.correlationId,
      RequestedById: me.Id
    });
    return (result as { Id: number }).Id;
  }

  public async updateJobStatus(itemId: number, status: JobStatus): Promise<void> {
    await this._sp.web.lists
      .getByTitle(LIST_PROVISIONING_JOBS)
      .items.getById(itemId)
      .update({ Status: status });
  }

  public async approveJob(itemId: number): Promise<void> {
    const me: { Id: number } = await this._sp.web.currentUser.select('Id')();
    await this._sp.web.lists
      .getByTitle(LIST_PROVISIONING_JOBS)
      .items.getById(itemId)
      .update({ Status: 'Approved', ApprovedById: me.Id });
  }

  /** Persisted after every step attempt so a closed tab resumes cleanly (Section 6). */
  public async updateJobSteps(itemId: number, steps: IJobStep[]): Promise<void> {
    await this._sp.web.lists
      .getByTitle(LIST_PROVISIONING_JOBS)
      .items.getById(itemId)
      .update({ StepsJson: JSON.stringify(steps) });
  }

  public async setJobTargetUser(itemId: number, targetUserId: string): Promise<void> {
    await this._sp.web.lists
      .getByTitle(LIST_PROVISIONING_JOBS)
      .items.getById(itemId)
      .update({ TargetUserId: targetUserId });
  }

  // ---- Audit ---------------------------------------------------------------

  public async addAuditEntry(entry: IAuditEntry): Promise<void> {
    await this._sp.web.lists.getByTitle(LIST_AUDIT_LOG).items.add({
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
    });
  }

  /** Audit trail for one job, oldest first (viewAudit surfaces in the drawer). */
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
    }[] = await this._sp.web.lists
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
      .top(top)();
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

  // ---- Roles ---------------------------------------------------------------

  public async getRoleDefinitions(): Promise<IRoleDefinition[]> {
    const items: { Title: AppRole; MemberGroupId: string; PermissionsJson: string | null }[] =
      await this._sp.web.lists
        .getByTitle(LIST_ROLES)
        .items.select('Title', 'MemberGroupId', 'PermissionsJson')
        .top(50)();
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

  /** Every UPC_Roles row, including ones with no MemberGroupId set yet — for the role-management UI. */
  public async getRoleDefinitionsForManagement(): Promise<IRoleManagementItem[]> {
    const items: { Id: number; Title: AppRole; MemberGroupId: string | null; PermissionsJson: string | null }[] =
      await this._sp.web.lists
        .getByTitle(LIST_ROLES)
        .items.select('Id', 'Title', 'MemberGroupId', 'PermissionsJson')
        .top(50)();
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

  public async updateRoleDefinition(
    itemId: number,
    memberGroupId: string,
    permissions: AppPermission[]
  ): Promise<void> {
    await this._sp.web.lists
      .getByTitle(LIST_ROLES)
      .items.getById(itemId)
      .update({ MemberGroupId: memberGroupId, PermissionsJson: JSON.stringify(permissions) });
  }

  // ---- Reference data --------------------------------------------------------

  public async getLicenseCosts(): Promise<ILicenseCost[]> {
    const items: { Title: string; MonthlyCost: number; Currency: string }[] = await this._sp.web.lists
      .getByTitle(LIST_LICENSE_COST_TABLE)
      .items.select('Title', 'MonthlyCost', 'Currency')
      .top(500)();
    return items.map((i) => ({
      skuPartNumber: i.Title,
      monthlyCost: i.MonthlyCost,
      currency: i.Currency
    }));
  }

  /** Curated Teams an operator/template can grant membership in. */
  public async getTeamsCatalog(): Promise<ITeamsCatalogItem[]> {
    const items: { Id: number; Title: string; TeamId: string; Category: string | null; DefaultRole: string | null }[] =
      await this._sp.web.lists
        .getByTitle(LIST_TEAMS_CATALOG)
        .items.select('Id', 'Title', 'TeamId', 'Category', 'DefaultRole')
        .top(500)();
    return items.map((i) => ({
      itemId: i.Id,
      title: i.Title,
      teamId: i.TeamId,
      category: i.Category ?? '',
      defaultRole: i.DefaultRole === 'owner' ? 'owner' : 'member'
    }));
  }

  /** Curated SharePoint sites an operator/template can grant access to. */
  public async getSiteCatalog(): Promise<ISiteCatalogItem[]> {
    const items: {
      Id: number;
      Title: string;
      SiteUrl: string;
      BusinessOwner?: { Title: string } | null;
      Category: string | null;
    }[] = await this._sp.web.lists
      .getByTitle(LIST_SITE_CATALOG)
      .items.select('Id', 'Title', 'SiteUrl', 'BusinessOwner/Title', 'Category')
      .expand('BusinessOwner')
      .top(500)();
    return items.map((i) => ({
      itemId: i.Id,
      title: i.Title,
      siteUrl: i.SiteUrl,
      businessOwner: i.BusinessOwner?.Title ?? null,
      category: i.Category ?? ''
    }));
  }

  /** Curated applications an operator/template can grant. */
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
    }[] = await this._sp.web.lists
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
      .top(500)();
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
      await this._sp.web.lists
        .getByTitle(LIST_DEPARTMENT_TEMPLATES)
        .items.select('Id', 'Title', 'TemplateJson', 'IsActive', 'Version')
        .filter('IsActive eq 1')
        .top(200)();
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
      } catch {
        // Malformed TemplateJson rows are skipped; template admin UI (Phase 2) surfaces them.
      }
    }
    return templates;
  }

  /** Admin view: every template row, including inactive and malformed ones. */
  public async getAllTemplates(): Promise<ITemplateListItem[]> {
    const items: { Id: number; Title: string; TemplateJson: string; IsActive: boolean; Version: number }[] =
      await this._sp.web.lists
        .getByTitle(LIST_DEPARTMENT_TEMPLATES)
        .items.select('Id', 'Title', 'TemplateJson', 'IsActive', 'Version')
        .top(200)();
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
      // Rows saved before approverGroupId existed have no such key in their
      // stored JSON — treat that the same as "no restriction" rather than
      // leaving it as undefined.
      template.approverGroupId = template.approverGroupId ?? null;
      return { itemId: i.Id, title: i.Title, template, isActive: i.IsActive, version: i.Version };
    });
  }

  public async createTemplate(title: string, template: IDepartmentTemplate): Promise<number> {
    const result = await this._sp.web.lists.getByTitle(LIST_DEPARTMENT_TEMPLATES).items.add({
      Title: title,
      TemplateJson: JSON.stringify(template),
      IsActive: true,
      Version: 1
    });
    return (result as { Id: number }).Id;
  }

  public async updateTemplate(
    itemId: number,
    title: string,
    template: IDepartmentTemplate,
    currentVersion: number
  ): Promise<void> {
    await this._sp.web.lists
      .getByTitle(LIST_DEPARTMENT_TEMPLATES)
      .items.getById(itemId)
      .update({
        Title: title,
        TemplateJson: JSON.stringify(template),
        Version: currentVersion + 1
      });
  }

  public async setTemplateActive(itemId: number, isActive: boolean): Promise<void> {
    await this._sp.web.lists
      .getByTitle(LIST_DEPARTMENT_TEMPLATES)
      .items.getById(itemId)
      .update({ IsActive: isActive });
  }

  // ---- Settings --------------------------------------------------------------

  /**
   * Tenant-shared app settings (single 'app' row in UPC_Settings). Returns a
   * partial: callers merge over DEFAULT_APP_SETTINGS so missing keys and
   * pre-upgrade rows behave sensibly.
   */
  public async getAppSettings(): Promise<Partial<IAppSettings>> {
    const items: { Id: number; SettingsJson: string | null }[] = await this._sp.web.lists
      .getByTitle(LIST_SETTINGS)
      .items.select('Id', 'SettingsJson')
      .filter("Title eq 'app'")
      .top(1)();
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
    const list = this._sp.web.lists.getByTitle(LIST_SETTINGS);
    const items: { Id: number }[] = await list.items
      .select('Id')
      .filter("Title eq 'app'")
      .top(1)();
    const payload = { Title: 'app', SettingsJson: JSON.stringify(settings) };
    if (items.length > 0) {
      await list.items.getById(items[0].Id).update(payload);
    } else {
      await list.items.add(payload);
    }
  }

  // ---- Tasks -----------------------------------------------------------------

  /** Service-desk queue, newest first. */
  public async getTasks(top: number = 500): Promise<IServiceDeskTask[]> {
    const items: {
      Id: number;
      Title: string;
      JobId: string | null;
      TaskType: TaskType;
      Instructions: string | null;
      Status: TaskStatus;
      CompletedBy: string | null;
      CompletedUtc: string | null;
      AssignedTo?: { Title: string } | null;
    }[] = await this._sp.web.lists
      .getByTitle(LIST_TASKS)
      .items.select(
        'Id',
        'Title',
        'JobId',
        'TaskType',
        'Instructions',
        'Status',
        'CompletedBy',
        'CompletedUtc',
        'AssignedTo/Title'
      )
      .expand('AssignedTo')
      .orderBy('Id', false)
      .top(top)();
    return items.map((i) => ({
      itemId: i.Id,
      title: i.Title,
      jobId: i.JobId ?? '',
      taskType: i.TaskType,
      instructions: i.Instructions ?? '',
      assignedTo: i.AssignedTo?.Title ?? null,
      status: i.Status,
      completedBy: i.CompletedBy ?? null,
      completedUtc: i.CompletedUtc ?? null
    }));
  }

  public async completeTask(itemId: number): Promise<void> {
    const me: { Title: string } = await this._sp.web.currentUser.select('Title')();
    await this._sp.web.lists
      .getByTitle(LIST_TASKS)
      .items.getById(itemId)
      .update({
        Status: 'Done',
        CompletedBy: me.Title,
        CompletedUtc: new Date().toISOString()
      });
  }

  /** Landing place for operations Graph cannot perform (spec Section 10). */
  public async createTask(
    jobId: string,
    taskType: TaskType,
    title: string,
    instructions: string
  ): Promise<void> {
    await this._sp.web.lists.getByTitle(LIST_TASKS).items.add({
      Title: title,
      JobId: jobId,
      TaskType: taskType,
      Instructions: instructions,
      Status: 'Open'
    });
  }
}
