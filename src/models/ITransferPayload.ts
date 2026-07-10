import type { ILicenseSelection } from './IOnboardingPayload';

export interface ITransferTarget {
  userId: string;
  displayName: string;
  userPrincipalName: string;
}

export interface ITransferChanges {
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  /** Entra objectId of the new manager; undefined/empty leaves it unchanged. */
  managerId?: string;
  managerDisplayName?: string;
  /** SKUs to add on top of whatever the user already holds. */
  addLicenses: ILicenseSelection[];
  /** skuIds to remove from the user's current assignment. */
  removeLicenseSkuIds: string[];
}

/**
 * Transfer job payload — changes department, job title, office, manager
 * and/or licenses for an EXISTING user (unlike onboarding/clone, which
 * create one). Every field in `changes` is optional; only set fields are
 * applied, so a transfer that only changes department leaves everything
 * else untouched.
 */
export interface ITransferPayload {
  schemaVersion: 1;
  kind: 'transfer';
  target: ITransferTarget;
  changes: ITransferChanges;
}
