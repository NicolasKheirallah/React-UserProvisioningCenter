export interface IAppSettings {
  requireApproval: boolean;
  requiredApprovals: number;
  bulkRowLimit: number;
  jobsRefreshSeconds: number;
}

export const DEFAULT_APP_SETTINGS: IAppSettings = {
  requireApproval: true,
  requiredApprovals: 1,
  bulkRowLimit: 100,
  jobsRefreshSeconds: 30
};
