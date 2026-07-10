import { validatePayload } from './payloadValidator';
import type { IOnboardingPayload } from '../models';

function validPayload(): IOnboardingPayload {
  return {
    schemaVersion: 1,
    kind: 'onboard',
    personal: {
      firstName: 'Anna',
      lastName: 'Svensson',
      displayName: 'Anna Svensson',
      employeeId: 'E12345'
    },
    employment: {
      jobTitle: 'Controller',
      department: 'Finance',
      employeeType: 'Employee',
      hireDate: '2026-08-01',
      managerId: '11111111-2222-3333-4444-555555555555'
    },
    identity: {
      userPrincipalName: 'anna.svensson@contoso.com',
      mailNickname: 'anna.svensson',
      domain: 'contoso.com',
      accountType: 'member'
    },
    accountSettings: {
      usageLocation: 'SE',
      accountEnabled: true,
      credentialMode: 'tap',
      forceChangePassword: true
    },
    licenses: [{ skuId: 'abc', skuPartNumber: 'SPE_E5' }],
    access: {
      securityGroups: [],
      m365Groups: [],
      teams: [],
      sharePointSites: [],
      applications: []
    },
    expirationReviewDays: null
  };
}

describe('validatePayload', () => {
  it('accepts a valid payload', () => {
    expect(validatePayload(validPayload())).toEqual([]);
  });

  it('rejects unknown schema versions', () => {
    const p = validPayload();
    (p as unknown as { schemaVersion: number }).schemaVersion = 2;
    expect(validatePayload(p)[0]).toContain('schema version');
  });

  it('requires employeeId', () => {
    const p = validPayload();
    p.personal.employeeId = ' ';
    expect(validatePayload(p)).toContainEqual(expect.stringContaining('employeeId'));
  });

  it('rejects UPN that does not match local@domain', () => {
    const p = validPayload();
    p.identity.userPrincipalName = 'other@contoso.com';
    expect(validatePayload(p)).toContainEqual(expect.stringContaining('userPrincipalName'));
  });

  it('rejects a local part with a trailing dot', () => {
    const p = validPayload();
    p.identity.mailNickname = 'anna.svensson.';
    p.identity.userPrincipalName = 'anna.svensson.@contoso.com';
    expect(validatePayload(p)).toContainEqual(expect.stringContaining('mailNickname'));
  });

  it('rejects invalid usage location', () => {
    const p = validPayload();
    p.accountSettings.usageLocation = 'Sweden';
    expect(validatePayload(p)).toContainEqual(expect.stringContaining('usageLocation'));
  });

  it('rejects duplicate license SKUs', () => {
    const p = validPayload();
    p.licenses.push({ skuId: 'abc', skuPartNumber: 'SPE_E5' });
    expect(validatePayload(p)).toContainEqual(expect.stringContaining('duplicate skuId'));
  });

  it('rejects invalid hire dates', () => {
    const p = validPayload();
    p.employment.hireDate = '01/08/2026';
    expect(validatePayload(p)).toContainEqual(expect.stringContaining('hireDate'));
  });
});
