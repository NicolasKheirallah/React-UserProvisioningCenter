import type { IIdentityInfo } from './IOnboardingPayload';

/** What happens to the mailbox when the user leaves. */
export type MailboxAction = 'none' | 'convertShared' | 'forward';

export interface IOffboardingOptions {
  /** Remove all directly assigned licenses. */
  removeLicenses: boolean;
  /** Remove from all groups the directory allows Graph to modify. */
  removeFromGroups: boolean;
  mailboxAction: MailboxAction;
  /** Required when mailboxAction is 'forward'. */
  forwardingAddress?: string;
  /** UPN that receives OneDrive access (usually the manager); creates a service-desk task. */
  oneDriveAccessUpn?: string;
}

export interface IOffboardingTarget {
  /** Entra objectId of the user being offboarded. */
  userId: string;
  displayName: string;
  userPrincipalName: string;
}

/**
 * Offboarding job payload. `identity` mirrors the onboarding shape so shared
 * surfaces (job list target column, audit, drawer meta) work on either payload.
 * Sign-in block + session revocation are always performed and are not options.
 */
export interface IOffboardingPayload {
  schemaVersion: 1;
  kind: 'offboard';
  identity: IIdentityInfo;
  target: IOffboardingTarget;
  options: IOffboardingOptions;
}
