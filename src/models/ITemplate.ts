/**
 * TemplateJson schema for UPC_DepartmentTemplates (spec Section 4).
 * Consumed by wizard steps 6-9 in Phase 2; defined now because the
 * provisioning script creates the list in Phase 1.
 */
export interface ITemplateLicense {
  skuPartNumber: string;
  required: boolean;
}

export interface ITemplateTeam {
  teamId: string;
  role: 'member' | 'owner';
}

export interface ITemplateSite {
  siteUrl: string;
  role: 'visitor' | 'member' | 'owner';
}

export interface IDepartmentTemplate {
  department: string;
  licenses: ITemplateLicense[];
  securityGroups: string[];
  m365Groups: string[];
  teams: ITemplateTeam[];
  sharePointSites: ITemplateSite[];
  applications: string[];
  /**
   * Entra security group object id. When set, a job created from this
   * template can only be approved by a member of this group (in addition
   * to holding the `approveJobs` permission) — see RoleService.isMemberOfGroup
   * and JobDetailDrawer's showApprove gate. Null/empty means no additional
   * restriction: any operator with approveJobs can approve it.
   */
  approverGroupId: string | null;
  expirationPolicyDays: number | null;
  usageLocationDefault: string;
}

/** Parsed UPC_DepartmentTemplates row. */
export interface ITemplateListItem {
  itemId: number;
  title: string;
  template: IDepartmentTemplate;
  isActive: boolean;
  version: number;
}
