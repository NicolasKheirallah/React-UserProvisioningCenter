export type AppRole =
  | 'ITAdmin'
  | 'HRAdmin'
  | 'DepartmentManager'
  | 'ServiceDesk'
  | 'Auditor'
  | 'ReadOnly';

/**
 * App-level permission verbs referenced by PermissionsJson.
 * These drive UI visibility ONLY — effective access is always the signed-in
 * operator's own Entra directory role (spec Section 1).
 */
export type AppPermission =
  | 'createJobs'
  | 'approveJobs'
  | 'runJobs'
  | 'retrySteps'
  | 'skipSteps'
  | 'cancelJobs'
  | 'manageTemplates'
  | 'viewAudit'
  | 'manageTasks'
  | 'manageSettings';

/** Parsed UPC_Roles row. */
export interface IRoleDefinition {
  role: AppRole;
  memberGroupId: string;
  permissions: AppPermission[];
}

/**
 * UPC_Roles row for the role-management UI — unlike IRoleDefinition (used by
 * RoleService's runtime resolution, which drops rows with no MemberGroupId
 * yet), this includes every row and the SharePoint item id needed to save
 * edits back.
 */
export interface IRoleManagementItem extends IRoleDefinition {
  itemId: number;
}

/** Resolved roles for the current operator (cached 15 minutes per session). */
export interface IResolvedRoles {
  roles: AppRole[];
  permissions: Set<AppPermission>;
  resolvedUtc: string;
}
