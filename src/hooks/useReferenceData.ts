import { useQuery, type UseQueryResult } from '@tanstack/react-query';
import { useServices } from '../contexts/ServicesContext';
import { useSettings } from '../contexts/SettingsContext';
import {
  QK_APP_ROLES,
  QK_APPLICATION_CATALOG,
  QK_JOBS,
  QK_PREFLIGHT,
  QK_ROLES,
  QK_SITE_CATALOG,
  QK_SUBSCRIBED_SKUS,
  QK_TASKS,
  QK_TEAMS_CATALOG,
  QK_TEMPLATES,
  QK_TEMPLATES_ALL,
  QK_VERIFIED_DOMAINS
} from '../constants/queryKeys';
import type {
  IApplicationCatalogItem,
  ILicenseOption,
  IPreflightResult,
  IProvisioningJob,
  IResolvedRoles,
  IRoleManagementItem,
  IServiceDeskTask,
  ISiteCatalogItem,
  ITeamsCatalogItem,
  ITemplateListItem
} from '../models';
import type { IPagedResult } from '../services/sharePointData/SharePointDataService';
import type { IVerifiedDomain } from '../services/users/UserService';

const STATIC_STALE_MS: number = 15 * 60 * 1000;
/** Shorter staleness window for admin-facing but operationally-live lists. */
const OPERATIONAL_STALE_MS: number = 60 * 1000;

export function usePreflight(): UseQueryResult<IPreflightResult> {
  const services = useServices();
  return useQuery(QK_PREFLIGHT, ({ signal }) => services.preflight.run(signal), {
    staleTime: STATIC_STALE_MS,
    retry: false
  });
}

export function useAppRoles(): UseQueryResult<IResolvedRoles> {
  const services = useServices();
  // forceRefresh left false so RoleService's own 15-minute cache (shared
  // across any non-React-Query callers) and this query's staleTime don't
  // fight each other — RoleService.getResolvedRoles is idempotent within
  // its TTL either way.
  return useQuery(QK_APP_ROLES, () => services.roles.getResolvedRoles(), {
    staleTime: STATIC_STALE_MS
  });
}

export function useLicenseOptions(): UseQueryResult<ILicenseOption[]> {
  const services = useServices();
  return useQuery(QK_SUBSCRIBED_SKUS, ({ signal }) => services.licenses.getLicenseOptions(signal), {
    staleTime: STATIC_STALE_MS
  });
}

export function useVerifiedDomains(): UseQueryResult<IVerifiedDomain[]> {
  const services = useServices();
  return useQuery(QK_VERIFIED_DOMAINS, ({ signal }) => services.users.getVerifiedDomains(signal), {
    staleTime: STATIC_STALE_MS
  });
}

/** Active templates for the wizard's "start from template" picker. */
export function useActiveTemplates(): UseQueryResult<ITemplateListItem[]> {
  const services = useServices();
  return useQuery(QK_TEMPLATES, () => services.data.getActiveTemplates(), {
    staleTime: STATIC_STALE_MS
  });
}

/** All templates (including inactive) for the management tab. */
export function useAllTemplates(): UseQueryResult<ITemplateListItem[]> {
  const services = useServices();
  return useQuery(QK_TEMPLATES_ALL, () => services.data.getAllTemplates(), {
    staleTime: STATIC_STALE_MS
  });
}

/** Service-desk task queue (Tasks tab). */
export function useTasks(): UseQueryResult<IServiceDeskTask[]> {
  const services = useServices();
  return useQuery(QK_TASKS, () => services.data.getTasks(), {
    staleTime: OPERATIONAL_STALE_MS
  });
}

/** Curated Teams catalog for the Access wizard step and template editor. */
export function useTeamsCatalog(): UseQueryResult<ITeamsCatalogItem[]> {
  const services = useServices();
  return useQuery(QK_TEAMS_CATALOG, () => services.data.getTeamsCatalog(), {
    staleTime: STATIC_STALE_MS
  });
}

/** Curated SharePoint site catalog for the Access wizard step and template editor. */
export function useSiteCatalog(): UseQueryResult<ISiteCatalogItem[]> {
  const services = useServices();
  return useQuery(QK_SITE_CATALOG, () => services.data.getSiteCatalog(), {
    staleTime: STATIC_STALE_MS
  });
}

/** Curated application catalog for the Access wizard step and template editor. */
export function useApplicationCatalog(): UseQueryResult<IApplicationCatalogItem[]> {
  const services = useServices();
  return useQuery(QK_APPLICATION_CATALOG, () => services.data.getApplicationCatalog(), {
    staleTime: STATIC_STALE_MS
  });
}

/** Every UPC_Roles row (including unconfigured ones) for the role-management UI. */
export function useRoleDefinitionsForManagement(): UseQueryResult<IRoleManagementItem[]> {
  const services = useServices();
  return useQuery(QK_ROLES, () => services.data.getRoleDefinitionsForManagement(), {
    staleTime: OPERATIONAL_STALE_MS
  });
}

export function useJobs(): UseQueryResult<IPagedResult<IProvisioningJob>> {
  const services = useServices();
  const { jobsRefreshSeconds } = useSettings();
  return useQuery(QK_JOBS, () => services.data.getJobsPaged(), {
    // Poll fast while something is executing so the dashboard reads as live,
    // fall back to the configured idle heartbeat (Settings tab).
    refetchInterval: (data) =>
      (data?.items ?? []).some((j) => j.status === 'Running') ? 5 * 1000 : jobsRefreshSeconds * 1000
  });
}
