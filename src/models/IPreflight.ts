export type CapabilityId =
  | 'directoryRead'
  | 'licenseRead'
  | 'sharePointRead'
  | 'sharePointWrite'
  | 'groupMemberRead'
  | 'createUsers'
  | 'assignLicenses'
  | 'groupWrites'
  | 'teamsWrites'
  | 'tapCreation'
  | 'guestInvites'
  | 'revokeSessions';

export interface ICapabilityCheck {
  capability: CapabilityId;
  label: string;
  ok: boolean;
  detail: string;
}

export interface IPreflightResult {
  checks: ICapabilityCheck[];
  missing: ICapabilityCheck[];
  directoryRoleTemplateIds: string[];
  operatorUpn: string;
  requiredGraphScopes: string[];
}
