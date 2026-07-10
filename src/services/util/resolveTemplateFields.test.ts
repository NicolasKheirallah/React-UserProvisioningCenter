import { resolveTemplateFields } from './resolveTemplateFields';
import type { IDepartmentTemplate, ILicenseOption } from '../../models';

function template(overrides?: Partial<IDepartmentTemplate>): IDepartmentTemplate {
  return {
    department: 'Finance',
    licenses: [{ skuPartNumber: 'SPE_E5', required: false }],
    securityGroups: ['grp-1'],
    m365Groups: ['grp-2'],
    teams: [{ teamId: 'team-1', role: 'member' }],
    sharePointSites: [{ siteUrl: 'https://contoso.sharepoint.com/sites/finance', role: 'member' }],
    applications: ['1'],
    approverGroupId: null,
    expirationPolicyDays: 90,
    usageLocationDefault: 'SE',
    ...overrides
  };
}

function licenseOptions(): ILicenseOption[] {
  return [
    {
      skuId: 'sku-guid-1',
      skuPartNumber: 'SPE_E5',
      displayName: 'Microsoft 365 E5',
      enabledUnits: 50,
      consumedUnits: 40,
      availableUnits: 10,
      capabilityStatus: 'Enabled',
      includesExchange: true,
      monthlyCost: null,
      currency: null
    }
  ];
}

describe('resolveTemplateFields', () => {
  it('resolves the template license skuPartNumber to this tenant\'s skuId', () => {
    const result = resolveTemplateFields(template(), licenseOptions());
    expect(result.licenses).toEqual([
      { skuId: 'sku-guid-1', skuPartNumber: 'SPE_E5', displayName: 'Microsoft 365 E5' }
    ]);
  });

  it('drops a template license with no matching tenant SKU', () => {
    const result = resolveTemplateFields(template(), []);
    expect(result.licenses).toEqual([]);
  });

  it('carries department, usage location, access grants, and expiration through unchanged', () => {
    const result = resolveTemplateFields(template(), licenseOptions());
    expect(result.department).toBe('Finance');
    expect(result.usageLocation).toBe('SE');
    expect(result.access).toEqual({
      securityGroups: ['grp-1'],
      m365Groups: ['grp-2'],
      teams: [{ teamId: 'team-1', role: 'member' }],
      sharePointSites: [{ siteUrl: 'https://contoso.sharepoint.com/sites/finance', role: 'member' }],
      applications: ['1']
    });
    expect(result.expirationReviewDays).toBe(90);
  });

  it('maps a blank usageLocationDefault to undefined rather than an empty string', () => {
    const result = resolveTemplateFields(template({ usageLocationDefault: '' }), licenseOptions());
    expect(result.usageLocation).toBeUndefined();
  });
});
