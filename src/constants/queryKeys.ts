/** TanStack Query cache keys. Static reference data gets long staleTimes. */
export const QK_SUBSCRIBED_SKUS: string[] = ['graph', 'subscribedSkus'];
export const QK_LICENSE_COSTS: string[] = ['sp', 'licenseCosts'];
export const QK_VERIFIED_DOMAINS: string[] = ['graph', 'verifiedDomains'];
export const QK_APP_ROLES: string[] = ['app', 'roles'];
export const QK_PREFLIGHT: string[] = ['app', 'preflight'];
export const QK_JOBS: string[] = ['sp', 'jobs'];
export const QK_TEMPLATES: string[] = ['sp', 'templates'];
export const QK_TEMPLATES_ALL: string[] = ['sp', 'templates', 'all'];
export const QK_TASKS: string[] = ['sp', 'tasks'];
export const QK_SETTINGS: string[] = ['sp', 'settings'];
/** Per-job audit entries: use [...QK_AUDIT, jobId]. */
export const QK_AUDIT: string[] = ['sp', 'audit'];
export const QK_TEAMS_CATALOG: string[] = ['sp', 'teamsCatalog'];
export const QK_SITE_CATALOG: string[] = ['sp', 'siteCatalog'];
export const QK_APPLICATION_CATALOG: string[] = ['sp', 'applicationCatalog'];
export const QK_ROLES: string[] = ['sp', 'roleDefinitions'];
