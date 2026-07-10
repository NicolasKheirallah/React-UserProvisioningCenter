/** Canonical step identifiers, in onboarding execution order. */
export const STEP_VALIDATE_INPUT: string = 'validate-input';
export const STEP_CREATE_USER: string = 'create-user';
export const STEP_SET_USAGE_LOCATION: string = 'set-usage-location';
export const STEP_ASSIGN_MANAGER: string = 'assign-manager';
export const STEP_ASSIGN_LICENSES: string = 'assign-licenses';
export const STEP_ASSIGN_GROUPS: string = 'assign-groups';
export const STEP_ASSIGN_TEAMS: string = 'assign-teams';
export const STEP_ASSIGN_SHAREPOINT: string = 'assign-sharepoint';
export const STEP_ASSIGN_APPLICATIONS: string = 'assign-applications';
export const STEP_UPLOAD_PHOTO: string = 'upload-photo';
export const STEP_PRESENT_CREDENTIALS: string = 'present-credentials';
export const STEP_SEND_NOTIFICATIONS: string = 'send-notifications';
export const STEP_SCHEDULE_ACCESS_REVIEW: string = 'schedule-access-review';
export const STEP_FINALIZE_AUDIT: string = 'finalize-audit';

/** Offboarding step identifiers, in execution order. */
export const STEP_VALIDATE_TARGET: string = 'validate-target';
export const STEP_BLOCK_SIGN_IN: string = 'block-sign-in';
export const STEP_REMOVE_LICENSES: string = 'remove-licenses';
export const STEP_REMOVE_GROUPS: string = 'remove-groups';
export const STEP_CREATE_HANDOVER_TASKS: string = 'create-handover-tasks';

/** Transfer step identifiers (department/manager/license change), in execution order. */
export const STEP_VALIDATE_TRANSFER: string = 'validate-transfer';
export const STEP_UPDATE_EMPLOYMENT: string = 'update-employment';
export const STEP_UPDATE_MANAGER: string = 'update-manager';
export const STEP_UPDATE_LICENSES: string = 'update-licenses';

/** Clone step identifiers (duplicate an existing user's access profile), in execution order. */
export const STEP_VALIDATE_CLONE_SOURCE: string = 'validate-clone-source';
export const STEP_COPY_LICENSES: string = 'copy-licenses';
export const STEP_COPY_GROUPS: string = 'copy-groups';

export const ONBOARDING_STEP_ORDER: readonly string[] = [
  STEP_VALIDATE_INPUT,
  STEP_CREATE_USER,
  STEP_SET_USAGE_LOCATION,
  STEP_ASSIGN_MANAGER,
  STEP_ASSIGN_LICENSES,
  STEP_ASSIGN_GROUPS,
  STEP_ASSIGN_TEAMS,
  STEP_ASSIGN_SHAREPOINT,
  STEP_ASSIGN_APPLICATIONS,
  STEP_PRESENT_CREDENTIALS,
  STEP_SEND_NOTIFICATIONS,
  STEP_SCHEDULE_ACCESS_REVIEW,
  STEP_FINALIZE_AUDIT
];

export const OFFBOARDING_STEP_ORDER: readonly string[] = [
  STEP_VALIDATE_TARGET,
  STEP_BLOCK_SIGN_IN,
  STEP_REMOVE_LICENSES,
  STEP_REMOVE_GROUPS,
  STEP_CREATE_HANDOVER_TASKS,
  STEP_SEND_NOTIFICATIONS,
  STEP_FINALIZE_AUDIT
];

export const TRANSFER_STEP_ORDER: readonly string[] = [
  STEP_VALIDATE_TRANSFER,
  STEP_UPDATE_EMPLOYMENT,
  STEP_UPDATE_MANAGER,
  STEP_UPDATE_LICENSES,
  STEP_SEND_NOTIFICATIONS,
  STEP_FINALIZE_AUDIT
];

/**
 * Clone reuses the full onboarding pipeline (the cloned user is still a new
 * hire — needs create-user, usage location, its own template selections,
 * credentials, notifications) and inserts the copy-from-source steps right
 * after assign-licenses, before the new user's own group/Team/site/app
 * grants run, so those steps' idempotency checks see the copied ones too.
 */
export const CLONE_STEP_ORDER: readonly string[] = [
  STEP_VALIDATE_INPUT,
  STEP_CREATE_USER,
  STEP_SET_USAGE_LOCATION,
  STEP_ASSIGN_MANAGER,
  STEP_ASSIGN_LICENSES,
  STEP_VALIDATE_CLONE_SOURCE,
  STEP_COPY_LICENSES,
  STEP_COPY_GROUPS,
  STEP_ASSIGN_GROUPS,
  STEP_ASSIGN_TEAMS,
  STEP_ASSIGN_SHAREPOINT,
  STEP_ASSIGN_APPLICATIONS,
  STEP_UPLOAD_PHOTO,
  STEP_PRESENT_CREDENTIALS,
  STEP_SEND_NOTIFICATIONS,
  STEP_SCHEDULE_ACCESS_REVIEW,
  STEP_FINALIZE_AUDIT
];
