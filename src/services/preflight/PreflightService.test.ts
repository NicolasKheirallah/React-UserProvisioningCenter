import { PreflightService, REQUIRED_GRAPH_SCOPES } from './PreflightService';
import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { ICapabilityCheck } from '../../models';
import { GraphServiceError } from '../graph/GraphError';

describe('PreflightService', () => {
  function makeGraph(): {
    graph: GraphService;
    getHandlers: { [path: string]: () => unknown };
    postHandlers: { [path: string]: (body: unknown) => unknown };
    batchHandlers: { [key: string]: () => { status: number; body: unknown } };
  } {
    const getHandlers: { [path: string]: () => unknown } = {};
    const postHandlers: { [path: string]: (body: unknown) => unknown } = {};
    const batchHandlers: { [key: string]: () => { status: number; body: unknown } } = {};

    const graph = {
      get: async <T>(path: string): Promise<T> => {
        const handler = getHandlers[path];
        if (!handler) throw new GraphServiceError('not found', 404, 'NotFound', 'req');
        return handler() as T;
      },
      post: async <T>(path: string, body: unknown): Promise<T> => {
        const handler = postHandlers[path];
        if (!handler) throw new GraphServiceError('not found', 404, 'NotFound', 'req');
        return handler(body) as T;
      },
      batch: async (requests: { id: string; url: string; method: string }[]) => {
        const map = new Map<string, { status: number; body: unknown }>();
        for (const req of requests) {
          const handler = batchHandlers[req.id];
          map.set(req.id, handler ? handler() : { status: 404, body: {} });
        }
        return map;
      }
    } as unknown as GraphService;

    return { graph, getHandlers, postHandlers, batchHandlers };
  }

  function makeData(overrides?: {
    getRoleDefinitions?: () => Promise<unknown[]>;
    probeWriteAccess?: () => Promise<boolean>;
  }): SharePointDataService {
    return {
      getRoleDefinitions: overrides?.getRoleDefinitions ?? (async () => []),
      probeWriteAccess: overrides?.probeWriteAccess ?? (async () => true)
    } as unknown as SharePointDataService;
  }

  it('returns all-clear when every probe succeeds and operator is Global Admin', async () => {
    const { graph, getHandlers, postHandlers, batchHandlers } = makeGraph();
    getHandlers['/me?$select=userPrincipalName'] = () => ({ userPrincipalName: 'admin@contoso.com' });
    getHandlers['/me/memberOf/microsoft.graph.directoryRole?$select=displayName,roleTemplateId'] = () => ({
      value: [{ roleTemplateId: '62e90394-69f5-4237-9190-012177145e10' }]
    });
    batchHandlers.org = () => ({ status: 200, body: { id: 'tenant-id' } });
    batchHandlers.skus = () => ({ status: 200, body: { value: [] } });
    batchHandlers.users = () => ({ status: 200, body: { value: [{}] } });
    postHandlers['/me/checkMemberGroups'] = () => ({ value: [] });

    const service = new PreflightService(graph, makeData());
    const result = await service.run();

    expect(result.missing).toHaveLength(0);
    expect(result.operatorUpn).toBe('admin@contoso.com');
    expect(result.requiredGraphScopes).toEqual(REQUIRED_GRAPH_SCOPES);
  });

  it('reports missing capabilities when probes fail and operator has no roles', async () => {
    const { graph, getHandlers, postHandlers, batchHandlers } = makeGraph();
    getHandlers['/me?$select=userPrincipalName'] = () => ({ userPrincipalName: 'user@contoso.com' });
    getHandlers['/me/memberOf/microsoft.graph.directoryRole?$select=displayName,roleTemplateId'] = () => ({
      value: []
    });
    batchHandlers.org = () => ({ status: 403, body: {} });
    batchHandlers.skus = () => ({ status: 403, body: {} });
    batchHandlers.users = () => ({ status: 403, body: {} });
    postHandlers['/me/checkMemberGroups'] = () => {
      throw new GraphServiceError('denied', 403, 'Authorization_RequestDenied', 'req');
    };

    const service = new PreflightService(
      graph,
      makeData({
        getRoleDefinitions: async () => {
          throw new Error('unreadable');
        },
        probeWriteAccess: async () => {
          throw new Error('unwritable');
        }
      })
    );
    const result = await service.run();

    const missing = new Set(result.missing.map((c: ICapabilityCheck) => c.capability));
    expect(missing.has('directoryRead')).toBe(true);
    expect(missing.has('licenseRead')).toBe(true);
    expect(missing.has('sharePointRead')).toBe(true);
    expect(missing.has('sharePointWrite')).toBe(true);
    expect(missing.has('groupMemberRead')).toBe(true);
    expect(missing.has('createUsers')).toBe(true);
    expect(missing.has('assignLicenses')).toBe(true);
  });

  it('flags groupMemberRead ok for non-403 errors (permission is consented)', async () => {
    const { graph, getHandlers, postHandlers, batchHandlers } = makeGraph();
    getHandlers['/me?$select=userPrincipalName'] = () => ({ userPrincipalName: 'user@contoso.com' });
    getHandlers['/me/memberOf/microsoft.graph.directoryRole?$select=displayName,roleTemplateId'] = () => ({
      value: []
    });
    batchHandlers.org = () => ({ status: 200, body: {} });
    batchHandlers.skus = () => ({ status: 200, body: {} });
    batchHandlers.users = () => ({ status: 200, body: {} });
    postHandlers['/me/checkMemberGroups'] = () => {
      throw new GraphServiceError('invalid group', 400, 'Request_BadRequest', 'req');
    };

    const service = new PreflightService(graph, makeData());
    const result = await service.run();

    const groupMemberRead = result.checks.find((c) => c.capability === 'groupMemberRead');
    expect(groupMemberRead?.ok).toBe(true);
  });
});
