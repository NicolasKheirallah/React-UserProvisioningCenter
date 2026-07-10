import type { ITemplateSite, ITemplateTeam } from './ITemplate';

/**
 * Resolved access to grant during onboarding — a snapshot taken at wizard
 * submission time (seeded from a department template via applyTemplate,
 * then editable), not a live reference back to the template. Field names
 * intentionally mirror IDepartmentTemplate's so mapping template -> payload
 * is a straight copy.
 */
export interface IAccessGrants {
  /** Entra security group object ids. */
  securityGroups: string[];
  /** Entra Microsoft 365 (unified) group object ids. */
  m365Groups: string[];
  teams: ITemplateTeam[];
  sharePointSites: ITemplateSite[];
  /** UPC_ApplicationCatalog item ids (see IApplicationCatalogItem.itemId). */
  applications: string[];
}

export const EMPTY_ACCESS_GRANTS: IAccessGrants = {
  securityGroups: [],
  m365Groups: [],
  teams: [],
  sharePointSites: [],
  applications: []
};

export function hasAnyAccessGrant(grants: IAccessGrants): boolean {
  return (
    grants.securityGroups.length > 0 ||
    grants.m365Groups.length > 0 ||
    grants.teams.length > 0 ||
    grants.sharePointSites.length > 0 ||
    grants.applications.length > 0
  );
}
