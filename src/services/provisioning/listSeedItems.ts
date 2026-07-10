import { DEFAULT_APP_SETTINGS, type IAppSettings } from '../../models';
import type { AppPermission, AppRole } from '../../models';
import { DEFAULT_ROLE_PERMISSIONS } from '../roles/RoleService';
import { LIST_ROLES, LIST_SETTINGS } from '../../constants/listNames';

/**
 * Seed items created the first time a list is provisioned. Idempotent: the
 * provisioning service checks by Title before inserting, so re-running after a
 * partial failure or an upgrade only adds what is still missing.
 *
 * `UPC_Settings` gets one row (Title='app') with the default settings JSON so
 * the Settings tab is editable immediately and the app sees explicit values
 * rather than falling back silently.
 *
 * `UPC_Roles` gets six rows (one per AppRole) with the canonical permission
 * sets pre-filled. The admin only needs to paste the Entra security group
 * object id into `MemberGroupId` for each role — no need to hand-write the
 * permission JSON.
 */

/** Shape of a seed item: field name → value (matches PnPJS add payload). */
export type SeedItem = Record<string, unknown>;

export interface ISeedDefinition {
  /** List title these items belong to. */
  listTitle: string;
  /** Field used to detect existing items (typically 'Title'). */
  identityField: string;
  /** Items to insert when missing. */
  items: SeedItem[];
}

const APP_SETTINGS_ROW: SeedItem = {
  Title: 'app',
  SettingsJson: JSON.stringify(DEFAULT_APP_SETTINGS satisfies IAppSettings)
};

function roleRow(role: AppRole): SeedItem {
  return {
    Title: role,
    MemberGroupId: '',
    PermissionsJson: JSON.stringify(DEFAULT_ROLE_PERMISSIONS[role] satisfies AppPermission[])
  };
}

const ROLE_ROWS: SeedItem[] = (['ITAdmin', 'HRAdmin', 'DepartmentManager', 'ServiceDesk', 'Auditor', 'ReadOnly'] as AppRole[]).map(
  roleRow
);

export const UPC_SEED_DEFINITIONS: ISeedDefinition[] = [
  {
    listTitle: LIST_SETTINGS,
    identityField: 'Title',
    items: [APP_SETTINGS_ROW]
  },
  {
    listTitle: LIST_ROLES,
    identityField: 'Title',
    items: ROLE_ROWS
  }
];