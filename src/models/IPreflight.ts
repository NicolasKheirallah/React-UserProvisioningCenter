export type CapabilityId =
  | 'directoryRead'
  | 'licenseRead'
  | 'createUsers'
  | 'assignLicenses'
  | 'groupWrites'
  | 'teamsWrites'
  | 'tapCreation'
  | 'guestInvites'
  | 'revokeSessions';

export interface ICapabilityCheck {
  capability: CapabilityId;
  /** Localized label shown in the preflight MessageBar. */
  label: string;
  ok: boolean;
  /** e.g. which Entra role would grant it. */
  detail: string;
}

export interface IPreflightResult {
  checks: ICapabilityCheck[];
  missing: ICapabilityCheck[];
  /** Operator's active Entra directory role template ids. */
  directoryRoleTemplateIds: string[];
  operatorUpn: string;
}
