import type { IAccessGrants, IDepartmentTemplate, ILicenseOption, ILicenseSelection } from '../../models';

export interface IResolvedTemplateFields {
  department: string;
  usageLocation: string | undefined;
  licenses: ILicenseSelection[];
  access: IAccessGrants;
  expirationReviewDays: number | null;
}

/**
 * Maps a department template's stored fields onto what an onboarding payload
 * needs — resolving license skuPartNumbers to this tenant's actual skuIds
 * (templates store the portable part number, not a tenant-specific id) and
 * carrying the access grants through as-is. Shared by the wizard's
 * PersonalStep (applyTemplate) and BulkImport (per-row `template` column) so
 * both onboarding paths treat a template the same way.
 */
export function resolveTemplateFields(
  template: IDepartmentTemplate,
  licenseOptions: ILicenseOption[]
): IResolvedTemplateFields {
  const bySkuPart: Map<string, string> = new Map(licenseOptions.map((o) => [o.skuPartNumber, o.skuId]));
  const licenses: ILicenseSelection[] = template.licenses
    .map((l) => ({
      skuId: bySkuPart.get(l.skuPartNumber) ?? '',
      skuPartNumber: l.skuPartNumber,
      displayName: licenseOptions.filter((o) => o.skuPartNumber === l.skuPartNumber)[0]?.displayName
    }))
    .filter((l) => !!l.skuId);
  return {
    department: template.department,
    usageLocation: template.usageLocationDefault || undefined,
    licenses,
    access: {
      securityGroups: template.securityGroups,
      m365Groups: template.m365Groups,
      teams: template.teams,
      sharePointSites: template.sharePointSites,
      applications: template.applications
    },
    expirationReviewDays: template.expirationPolicyDays
  };
}
