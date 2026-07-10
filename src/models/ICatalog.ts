/** Parsed UPC_TeamsCatalog row — curated list of Teams operators/templates can grant. */
export interface ITeamsCatalogItem {
  itemId: number;
  /** Friendly name (Title column). */
  title: string;
  teamId: string;
  category: string;
  defaultRole: 'member' | 'owner';
}

/** Parsed UPC_SiteCatalog row — curated list of SharePoint sites operators/templates can grant. */
export interface ISiteCatalogItem {
  itemId: number;
  /** Friendly name (Title column). */
  title: string;
  siteUrl: string;
  businessOwner: string | null;
  category: string;
}

export type ApplicationProvisioningType = 'Manual' | 'GroupBased';

/** Parsed UPC_ApplicationCatalog row — curated list of apps operators/templates can grant. */
export interface IApplicationCatalogItem {
  itemId: number;
  /** Friendly name (Title column). */
  title: string;
  owner: string | null;
  /** GroupBased: add the user to targetGroupId. Manual: always routes to UPC_Tasks. */
  provisioningType: ApplicationProvisioningType;
  targetGroupId: string | null;
  approvalRequired: boolean;
  instructions: string;
  isActive: boolean;
}
