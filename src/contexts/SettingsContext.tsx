import * as React from 'react';
import { useQuery, type UseQueryResult } from '@tanstack/react-query';
import { QK_SETTINGS } from '../constants/queryKeys';
import { DEFAULT_APP_SETTINGS, type IAppSettings } from '../models';
import { useServices } from './ServicesContext';

export type { IAppSettings };

const SettingsContext: React.Context<IAppSettings> =
  React.createContext<IAppSettings>(DEFAULT_APP_SETTINGS);

/** Raw settings query — the Settings tab uses this for editing/saving. */
export function useAppSettingsQuery(): UseQueryResult<Partial<IAppSettings>> {
  const services = useServices();
  return useQuery(QK_SETTINGS, () => services.data.getAppSettings(), {
    staleTime: 5 * 60 * 1000,
    // Missing list (not provisioned yet) must not break the app — consumers
    // fall back to defaults; only the Settings tab surfaces the error.
    retry: false
  });
}

/**
 * Tenant-shared configuration from UPC_Settings, merged over defaults.
 * While loading (or when the list is missing) the defaults apply, which are
 * the safe choices (approval gate on).
 */
export const SettingsProvider: React.FC<{ children?: React.ReactNode }> = ({ children }) => {
  const services = useServices();
  const query = useAppSettingsQuery();
  const value: IAppSettings = React.useMemo(
    () => ({ ...DEFAULT_APP_SETTINGS, ...(query.data ?? {}) }),
    [query.data]
  );
  // Push the resolved settings into the engine so steps read live config.
  React.useEffect(() => {
    services.engine.updateSettings(value);
  }, [services.engine, value]);
  return <SettingsContext.Provider value={value}>{children}</SettingsContext.Provider>;
};

export function useSettings(): IAppSettings {
  return React.useContext(SettingsContext);
}
