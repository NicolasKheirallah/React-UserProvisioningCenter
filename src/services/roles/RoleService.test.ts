import { RoleService, DEFAULT_ROLE_PERMISSIONS } from './RoleService';
import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { AppPermission, IRoleDefinition } from '../../models';

const ALL_PERMISSIONS: AppPermission[] = [
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
];

function makeGraph(checkMemberGroups: (groupIds: string[]) => string[]): GraphService {
  return {
    post: async (_path: string, body: { groupIds: string[] }) => ({
      value: checkMemberGroups(body.groupIds)
    })
  } as unknown as GraphService;
}

function makeData(definitions: IRoleDefinition[], canManageWeb: boolean = false): SharePointDataService {
  return {
    getRoleDefinitions: async () => definitions,
    canManageWeb: async () => canManageWeb
  } as unknown as SharePointDataService;
}

describe('DEFAULT_ROLE_PERMISSIONS', () => {
  it('ITAdmin (the seeded superuser role) is granted every permission the app defines', () => {

    for (const permission of ALL_PERMISSIONS) {
      expect(DEFAULT_ROLE_PERMISSIONS.ITAdmin).toContain(permission);
    }
  });

  it('ReadOnly is granted nothing', () => {
    expect(DEFAULT_ROLE_PERMISSIONS.ReadOnly).toEqual([]);
  });
});

describe('RoleService', () => {
  const itAdminDef: IRoleDefinition = {
    role: 'ITAdmin',
    memberGroupId: 'group-it-admin',
    permissions: []
  };
  const readOnlyDef: IRoleDefinition = {
    role: 'ReadOnly',
    memberGroupId: 'group-read-only',
    permissions: []
  };

  it('falls back to DEFAULT_ROLE_PERMISSIONS when a UPC_Roles row has empty PermissionsJson', async () => {
    const graph = makeGraph(() => ['group-it-admin']);
    const data = makeData([itAdminDef]);
    const svc = new RoleService(graph, data);

    const resolved = await svc.getResolvedRoles();
    expect(resolved.roles).toEqual(['ITAdmin']);
    expect(resolved.permissions.has('manageSettings')).toBe(true);
  });

  it('uses the row-specific PermissionsJson when present, ignoring the default set', async () => {
    const graph = makeGraph(() => ['group-it-admin']);
    const data = makeData([{ ...itAdminDef, permissions: ['viewAudit'] }]);
    const svc = new RoleService(graph, data);

    const resolved = await svc.getResolvedRoles();
    expect(Array.from(resolved.permissions)).toEqual(['viewAudit']);
  });

  it('defaults to ReadOnly when the operator matches no role group', async () => {
    const graph = makeGraph(() => []);
    const data = makeData([itAdminDef, readOnlyDef]);
    const svc = new RoleService(graph, data);

    const resolved = await svc.getResolvedRoles();
    expect(resolved.roles).toEqual(['ReadOnly']);
    expect(resolved.permissions.size).toBe(0);
  });

  it('defaults to ReadOnly when the roles list is missing/unreadable', async () => {
    const graph = makeGraph(() => ['group-it-admin']);
    const data = {
      getRoleDefinitions: async () => {
        throw new Error('list not found');
      }
    } as unknown as SharePointDataService;
    const svc = new RoleService(graph, data);

    const resolved = await svc.getResolvedRoles();
    expect(resolved.roles).toEqual(['ReadOnly']);
  });

  it('merges permissions across every matching role', async () => {
    const graph = makeGraph(() => ['group-it-admin', 'group-read-only']);
    const data = makeData([
      { ...itAdminDef, permissions: ['createJobs'] },
      { ...readOnlyDef, permissions: ['viewAudit'] }
    ]);
    const svc = new RoleService(graph, data);

    const resolved = await svc.getResolvedRoles();
    expect(resolved.roles.sort()).toEqual(['ITAdmin', 'ReadOnly']);
    expect(Array.from(resolved.permissions).sort()).toEqual(['createJobs', 'viewAudit']);
  });

  it('chunks checkMemberGroups calls at 20 group ids per request', async () => {
    const definitions: IRoleDefinition[] = Array.from({ length: 45 }, (_, i) => ({
      role: 'ReadOnly',
      memberGroupId: `group-${i}`,
      permissions: []
    }));
    const chunkSizes: number[] = [];
    const graph = makeGraph((groupIds) => {
      chunkSizes.push(groupIds.length);
      return [];
    });
    const svc = new RoleService(graph, makeData(definitions));
    await svc.getResolvedRoles();

    expect(chunkSizes).toEqual([20, 20, 5]);
  });

  it('continues resolving when one checkMemberGroups chunk fails', async () => {
    let call: number = 0;
    const graph = {
      post: async (): Promise<{ value: string[] }> => {
        call++;
        if (call === 1) {
          throw new Error('throttled');
        }
        return { value: ['group-it-admin'] };
      }
    } as unknown as GraphService;
    const definitions: IRoleDefinition[] = [
      ...Array.from({ length: 20 }, (_, i) => ({
        role: 'ReadOnly' as const,
        memberGroupId: `group-${i}`,
        permissions: []
      })),
      itAdminDef
    ];
    const svc = new RoleService(graph, makeData(definitions));

    const resolved = await svc.getResolvedRoles();
    expect(resolved.roles).toContain('ITAdmin');
  });

  it('caches the result for the TTL and only re-resolves on forceRefresh', async () => {
    let calls: number = 0;
    const graph = makeGraph(() => {
      calls++;
      return ['group-it-admin'];
    });
    const svc = new RoleService(graph, makeData([itAdminDef]));

    await svc.getResolvedRoles();
    await svc.getResolvedRoles();
    expect(calls).toBe(1);

    await svc.getResolvedRoles(true);
    expect(calls).toBe(2);
  });
});

describe('RoleService.isMemberOfGroup', () => {
  it('returns true when checkMemberGroups confirms membership', async () => {
    const graph = makeGraph((groupIds) => groupIds);
    const svc = new RoleService(graph, makeData([]));

    expect(await svc.isMemberOfGroup('grp-approvers')).toBe(true);
  });

  it('returns false when the operator is not a member', async () => {
    const graph = makeGraph(() => []);
    const svc = new RoleService(graph, makeData([]));

    expect(await svc.isMemberOfGroup('grp-approvers')).toBe(false);
  });

  it('returns false without calling Graph for an empty group id', async () => {
    let called: boolean = false;
    const graph = makeGraph(() => {
      called = true;
      return [];
    });
    const svc = new RoleService(graph, makeData([]));

    expect(await svc.isMemberOfGroup('')).toBe(false);
    expect(called).toBe(false);
  });

  it('fails closed (false) when the Graph call throws', async () => {
    const graph = {
      post: async () => {
        throw new Error('network error');
      }
    } as unknown as GraphService;
    const svc = new RoleService(graph, makeData([]));

    expect(await svc.isMemberOfGroup('grp-approvers')).toBe(false);
  });
});

describe('RoleService bootstrap for an unconfigured install', () => {
  it('reports unconfigured when no role is mapped to a group', async () => {
    const graph = makeGraph(() => []);
    const svc = new RoleService(graph, makeData([]));

    const resolved = await svc.getResolvedRoles();
    expect(resolved.unconfigured).toBe(true);
  });

  it('grants configuration-only permissions to someone who can manage the site', async () => {
    const graph = makeGraph(() => []);
    const svc = new RoleService(graph, makeData([], true));

    const resolved = await svc.getResolvedRoles();
    expect(resolved.bootstrapGranted).toBe(true);
    expect(resolved.permissions.has('manageSettings')).toBe(true);
    expect(resolved.permissions.has('manageDelegations')).toBe(true);
    expect(resolved.permissions.has('viewAudit')).toBe(true);
  });

  it('never grants job execution through the bootstrap path', async () => {
    const graph = makeGraph(() => []);
    const svc = new RoleService(graph, makeData([], true));

    const resolved = await svc.getResolvedRoles();
    expect(resolved.permissions.has('createJobs')).toBe(false);
    expect(resolved.permissions.has('runJobs')).toBe(false);
    expect(resolved.permissions.has('approveJobs')).toBe(false);
    expect(resolved.permissions.has('rollbackJobs')).toBe(false);
  });

  it('grants nothing to a non-owner on an unconfigured install', async () => {
    const graph = makeGraph(() => []);
    const svc = new RoleService(graph, makeData([], false));

    const resolved = await svc.getResolvedRoles();
    expect(resolved.bootstrapGranted).toBe(false);
    expect(resolved.permissions.size).toBe(0);
    expect(resolved.roles).toEqual(['ReadOnly']);
  });

  it('does not bootstrap once a role is mapped, even for a site owner', async () => {
    const graph = makeGraph(() => []);
    const definitions: IRoleDefinition[] = [
      { role: 'ITAdmin', memberGroupId: 'group-it-admin', permissions: [] }
    ];
    const svc = new RoleService(graph, makeData(definitions, true));

    const resolved = await svc.getResolvedRoles();
    expect(resolved.unconfigured).toBe(false);
    expect(resolved.bootstrapGranted).toBe(false);
    expect(resolved.permissions.size).toBe(0);
  });

  it('survives a failing site-permission probe', async () => {
    const graph = makeGraph(() => []);
    const data = {
      getRoleDefinitions: async () => [],
      canManageWeb: async () => {
        throw new Error('permission probe failed');
      }
    } as unknown as SharePointDataService;
    const svc = new RoleService(graph, data);

    const resolved = await svc.getResolvedRoles();
    expect(resolved.bootstrapGranted).toBe(false);
    expect(resolved.roles).toEqual(['ReadOnly']);
  });
});
