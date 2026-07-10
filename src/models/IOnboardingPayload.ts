import type { IAccessGrants } from './IAccessGrants';

export type EmployeeType = 'Employee' | 'Contractor';
export type CredentialMode = 'password' | 'tap';

/** Wizard step 1 — Personal. */
export interface IPersonalInfo {
  firstName: string;
  lastName: string;
  displayName: string;
  /** Entra employeeId; uniqueness enforced against the directory before leaving step 1. */
  employeeId: string;
  mobilePhone?: string;
  personalEmail?: string;
  /**
   * Optional profile photo as a data: URL, capped client-side (~50KB) so the
   * SharePoint Note field storing PayloadJson stays reasonably sized — this
   * is the one field in the payload that isn't small structured data.
   */
  photoDataUrl?: string;
}

/** Wizard step 2 — Employment. */
export interface IEmploymentInfo {
  jobTitle: string;
  department: string;
  companyName?: string;
  officeLocation?: string;
  /** ISO 3166-1 alpha-2 — the user's physical/office country (a Graph profile
   *  attribute, distinct from accountSettings.usageLocation which drives
   *  license-assignment eligibility and may legitimately differ, e.g. a
   *  remote hire licensed against the employer's usage location). */
  country?: string;
  /** Entra objectId of the manager, resolved via people search. */
  managerId?: string;
  managerDisplayName?: string;
  managerUpn?: string;
  employeeType: EmployeeType;
  /** ISO date (yyyy-MM-dd). */
  hireDate: string;
}

/** Wizard step 3 — Identity. */
export interface IIdentityInfo {
  /** Full UPN, e.g. anna.svensson@contoso.com. For a guest, the external
   *  email address the invitation is sent to (Entra assigns the actual
   *  #EXT# UPN on acceptance). */
  userPrincipalName: string;
  /** UPN local part, max 64 chars, no leading/trailing dot. Not meaningful
   *  for a guest — derived the same way but never sent to Graph. */
  mailNickname: string;
  domain: string;
  /**
   * 'member' (default) creates a normal cloud user via POST /users. 'guest'
   * invites an external user via POST /invitations instead — the account
   * settings, licenses and mailbox-wait steps don't apply and are skipped.
   */
  accountType: 'member' | 'guest';
}

/** Wizard step 4 — Account settings. */
export interface IAccountSettings {
  /** ISO 3166-1 alpha-2, required before license assignment. */
  usageLocation: string;
  accountEnabled: boolean;
  /**
   * How first-sign-in credentials are handed over. Values are generated
   * client-side, held in memory only and displayed once (Section 6).
   */
  credentialMode: CredentialMode;
  forceChangePassword: boolean;
}

/** Wizard step 5 — Licenses. */
export interface ILicenseSelection {
  skuId: string;
  skuPartNumber: string;
  /** Friendly product name, populated when the selection is made. */
  displayName?: string;
}

/**
 * Full wizard submission — stored as PayloadJson on the job.
 * MUST NEVER contain passwords, TAPs or any other secret.
 */
export interface IOnboardingPayload {
  schemaVersion: 1;
  kind: 'onboard';
  personal: IPersonalInfo;
  employment: IEmploymentInfo;
  identity: IIdentityInfo;
  accountSettings: IAccountSettings;
  licenses: ILicenseSelection[];
  /** Security/M365 groups, Teams, SharePoint sites and applications to grant. */
  access: IAccessGrants;
  /**
   * Days after which this account's access should be reviewed (from the
   * template's expirationPolicyDays, snapshotted at submission). Null means
   * no review is scheduled. The engine never auto-acts on this — it only
   * creates a future-dated UPC_Tasks reminder (see stepHelpers.ts).
   */
  expirationReviewDays: number | null;
  /**
   * When set, this is a Clone job (see IProvisioningJob.jobType): the new
   * hire's licenses and security/M365 group memberships are copied from
   * this existing user's directory profile in addition to whatever this
   * payload's own `licenses`/`access` already specify.
   */
  cloneSourceUserId?: string;
  cloneSourceDisplayName?: string;
  /**
   * Snapshotted from the applied template's approverGroupId at submission
   * time (survives the template being edited/deleted afterward). Null/unset
   * means no additional approval restriction beyond the approveJobs
   * permission — see RoleService.isMemberOfGroup.
   */
  approverGroupId?: string | null;
}
