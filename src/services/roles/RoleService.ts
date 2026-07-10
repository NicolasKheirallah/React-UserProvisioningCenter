import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { AppPermission, AppRole, IResolvedRoles, IRoleDefinition } from '../../models';

const CACHE_TTL_MS: number = 15 * 60 * 1000;
const CHECK_MEMBER_GROUPS_LIMIT: number = 20;

/**
 * Fallback permission sets used when a UPC_Roles row has no PermissionsJson.
 * Reminder: this whole layer only controls what the UI shows and offers.
 * Effective access is always the operator's own Entra directory role
 * (spec Section 1) — nothing here is a security boundary.
 */
export const DEFAULT_ROLE_PERMISSIONS: Record<AppRole, AppPermission[]> = {
  ITAdmin: [
    'createJobs',
    'approveJobs',
    'runJobs',
    'retrySteps',
    'skipSteps',
    'cancelJobs',
    'manageTemplates',
    'viewAudit',
    'manageTasks',
    'manageSettings'
  ],
  HRAdmin: ['createJobs', 'approveJobs', 'viewAudit'],
  DepartmentManager: ['createJobs'],
  ServiceDesk: ['createJobs', 'runJobs', 'retrySteps', 'manageTasks'],
  Auditor: ['viewAudit'],
  ReadOnly: []
};

/**
 * Resolves the operator's app roles from UPC_Roles via transitive Entra
 * group membership (/me/checkMemberGroups), cached 15 minutes per session.
 */
export class RoleService {
  private readonly _graph: GraphService;
  private readonly _data: SharePointDataService;
  private _cache: { value: IResolvedRoles; expiresAt: number } | undefined;

  public constructor(graph: GraphService, data: SharePointDataService) {
    this._graph = graph;
    this._data = data;
  }

  public async getResolvedRoles(forceRefresh: boolean = false): Promise<IResolvedRoles> {
    if (!forceRefresh && this._cache && this._cache.expiresAt > Date.now()) {
      return this._cache.value;
    }
    const value: IResolvedRoles = await this._resolve();
    this._cache = { value, expiresAt: Date.now() + CACHE_TTL_MS };
    return value;
  }

  /**
   * Is the signed-in operator a member of an arbitrary Entra group? Used to
   * gate approval on a template's approverGroupId — a narrower check than
   * getResolvedRoles, which only knows about UPC_Roles-registered groups.
   */
  public async isMemberOfGroup(groupId: string, signal?: AbortSignal): Promise<boolean> {
    if (!groupId) {
      return false;
    }
    try {
      const result: { value: string[] } = await this._graph.post<{ value: string[] }>(
        '/me/checkMemberGroups',
        { groupIds: [groupId] },
        { signal }
      );
      return (result.value ?? []).indexOf(groupId) !== -1;
    } catch {
      // Unreachable/denied → treat as "not a member" (fail closed on the UI
      // affordance; this is never the actual enforcement point — see
      // RoleService's class doc comment).
      return false;
    }
  }

  private async _resolve(): Promise<IResolvedRoles> {
    let definitions: IRoleDefinition[] = [];
    try {
      definitions = await this._data.getRoleDefinitions();
    } catch {
      // Roles list missing/unreadable → ReadOnly. UI-only consequence.
      definitions = [];
    }

    const groupIds: string[] = definitions.map((d) => d.memberGroupId);
    const memberOf: Set<string> = new Set();
    for (let i = 0; i < groupIds.length; i += CHECK_MEMBER_GROUPS_LIMIT) {
      const chunk: string[] = groupIds.slice(i, i + CHECK_MEMBER_GROUPS_LIMIT);
      try {
        const result: { value: string[] } = await this._graph.post<{ value: string[] }>(
          '/me/checkMemberGroups',
          { groupIds: chunk }
        );
        for (const id of result.value ?? []) {
          memberOf.add(id);
        }
      } catch {
        // A failed chunk only hides UI affordances; continue with the rest.
      }
    }

    const roles: AppRole[] = [];
    const permissions: Set<AppPermission> = new Set();
    for (const definition of definitions) {
      if (!memberOf.has(definition.memberGroupId)) {
        continue;
      }
      roles.push(definition.role);
      const effective: AppPermission[] =
        definition.permissions.length > 0
          ? definition.permissions
          : DEFAULT_ROLE_PERMISSIONS[definition.role] ?? [];
      for (const permission of effective) {
        permissions.add(permission);
      }
    }
    if (roles.length === 0) {
      roles.push('ReadOnly');
    }
    return { roles, permissions, resolvedUtc: new Date().toISOString() };
  }
}
