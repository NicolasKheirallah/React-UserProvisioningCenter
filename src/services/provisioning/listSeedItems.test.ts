import { UPC_SEED_DEFINITIONS } from './listSeedItems';
import { LIST_ROLES, LIST_SETTINGS } from '../../constants/listNames';
import { DEFAULT_APP_SETTINGS } from '../../models';
import { DEFAULT_ROLE_PERMISSIONS } from '../roles/RoleService';

describe('listSeedItems', () => {
  it('seeds UPC_Settings with one row titled app', () => {
    const settingsSeed = UPC_SEED_DEFINITIONS.find((s) => s.listTitle === LIST_SETTINGS);
    expect(settingsSeed).toBeDefined();
    expect(settingsSeed!.items).toHaveLength(1);
    expect(settingsSeed!.items[0].Title).toBe('app');
    const parsed = JSON.parse(settingsSeed!.items[0].SettingsJson as string);
    expect(parsed).toEqual(DEFAULT_APP_SETTINGS);
  });

  it('seeds UPC_Roles with six rows, one per AppRole', () => {
    const rolesSeed = UPC_SEED_DEFINITIONS.find((s) => s.listTitle === LIST_ROLES);
    expect(rolesSeed).toBeDefined();
    expect(rolesSeed!.items).toHaveLength(6);
    const titles = rolesSeed!.items.map((i) => i.Title);
    expect(titles).toEqual([
      'ITAdmin',
      'HRAdmin',
      'DepartmentManager',
      'ServiceDesk',
      'Auditor',
      'ReadOnly'
    ]);
  });

  it('each role row has PermissionsJson matching DEFAULT_ROLE_PERMISSIONS', () => {
    const rolesSeed = UPC_SEED_DEFINITIONS.find((s) => s.listTitle === LIST_ROLES)!;
    for (const item of rolesSeed.items) {
      const role = item.Title as keyof typeof DEFAULT_ROLE_PERMISSIONS;
      const permissions = JSON.parse(item.PermissionsJson as string);
      expect(permissions).toEqual(DEFAULT_ROLE_PERMISSIONS[role]);
    }
  });

  it('each role row has a blank MemberGroupId for the admin to fill', () => {
    const rolesSeed = UPC_SEED_DEFINITIONS.find((s) => s.listTitle === LIST_ROLES)!;
    for (const item of rolesSeed.items) {
      expect(item.MemberGroupId).toBe('');
    }
  });

  it('all seed definitions use Title as the identity field', () => {
    for (const seed of UPC_SEED_DEFINITIONS) {
      expect(seed.identityField).toBe('Title');
    }
  });
});