/**
 * Tenant-shared app configuration, stored as one JSON row in UPC_Settings
 * (Title = 'app') and edited from the Settings tab (manageSettings
 * permission). Unknown/missing keys fall back to the defaults so old rows
 * survive upgrades.
 */
export interface IAppSettings {
  /**
   * When true (default) new jobs start as PendingApproval and need an
   * approver before they can run; when false they are created ready to run.
   */
  requireApproval: boolean;
  /** Maximum rows accepted by one bulk CSV import. */
  bulkRowLimit: number;
  /** Dashboard idle refresh interval in seconds (5s while a job is running). */
  jobsRefreshSeconds: number;
}

export const DEFAULT_APP_SETTINGS: IAppSettings = {
  requireApproval: true,
  bulkRowLimit: 100,
  jobsRefreshSeconds: 30
};