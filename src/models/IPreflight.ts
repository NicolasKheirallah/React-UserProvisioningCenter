export type CapabilityId =
  | 'directoryRead'
  | 'licenseRead'
  | 'createUsers'
  | 'assignLicenses'
  | 'groupWrites'
  | 'teamsWrites'
  | 'tapCreation'
  | 'guestInvites'
  | 'revokeSessions'
  | 'sharePointRead'
  | 'sharePointWrite'
  | 'groupMemberRead';

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
