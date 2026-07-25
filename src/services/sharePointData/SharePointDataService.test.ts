jest.mock('@pnp/sp', () => ({ spfi: () => ({ using: () => ({}) }), SPFx: () => ({}), __esModule: true }));
jest.mock('@pnp/sp/webs', () => ({ __esModule: true }));
jest.mock('@pnp/sp/lists', () => ({ __esModule: true }));
jest.mock('@pnp/sp/items', () => ({ __esModule: true }));
jest.mock('@pnp/sp/site-users/web', () => ({ __esModule: true }));
jest.mock('@pnp/sp/security', () => ({ __esModule: true, PermissionKind: { AddListItems: 2 } }));

import { SharePointDataService } from './SharePointDataService';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

interface ICapture {
  select?: string[];
  expand?: string[];
  filter?: string;
  orderBy?: [string, boolean];
  top?: number;
}

function asyncIterableOf<T>(rows: T[], pageSize: number): AsyncIterable<T[]> {
  return {
    [Symbol.asyncIterator](): AsyncIterator<T[]> {
      let offset = 0;
      return {
        async next(): Promise<IteratorResult<T[]>> {
          if (offset >= rows.length && rows.length > 0) {
            return { value: undefined, done: true };
          }
          const page = rows.slice(offset, offset + pageSize);
          offset += pageSize;
          if (rows.length === 0) {
            return { value: [], done: false };
          }
          return { value: page, done: false };
        }
      };
    }
  };
}

function makeChain(returnValue: unknown[], capture: ICapture): unknown {
  const chain: Record<string, unknown> = {};
  chain.select = (...args: string[]): unknown => {
    capture.select = args;
    return chain;
  };
  chain.expand = (...args: string[]): unknown => {
    capture.expand = args;
    return chain;
  };
  chain.filter = (query: string): unknown => {
    capture.filter = query;
    return chain;
  };
  chain.orderBy = (field: string, asc: boolean): unknown => {
    capture.orderBy = [field, asc];
    return chain;
  };
  chain.top = (n: number): unknown => {
    capture.top = n;
    return chain;
  };
  chain.add = async (_payload: unknown): Promise<{ Id: number }> => ({ Id: 42 });
  chain.getById = (_id: number): unknown => ({
    ...chain,
    update: async (_payload: unknown): Promise<{ 'odata.etag': string }> => ({ 'odata.etag': '"1"' })
  });
  (chain as Record<PropertyKey, unknown>)[Symbol.asyncIterator] = (): AsyncIterator<unknown[]> =>
    asyncIterableOf(returnValue, capture.top ?? (returnValue.length || 1))[Symbol.asyncIterator]();
  return chain;
}

function makeMockSp(returnValue: unknown[], capture: ICapture, fieldNames?: string[]): unknown {
  const items = makeChain(returnValue, capture);
  const fields = {
    select: (..._args: string[]) => ({
      top: (_n: number) => async () => (fieldNames ?? []).map((InternalName) => ({ InternalName }))
    })
  };
  const lists = {
    getByTitle: (_title: string) => (fieldNames ? { items, fields } : { items }),
    ensure: async () => ({ created: true, list: {} })
  };
  const web = {
    lists,
    currentUser: {
      select: (..._fields: string[]) => async () => ({ Id: 1, Title: 'me', LoginName: 'me@contoso.com' })
    }
  };
  return { web, __root: 'mock' };
}

function makeJobRow(id: number): unknown {
  return {
    Id: id,
    Title: `job-${id}`,
    JobType: 'Onboard',
    Status: 'Completed',
    PayloadJson: '{}',
    StepsJson: '[]',
    ApprovalsJson: '[]',
    ScheduledFor: null,
    CorrelationId: 'c',
    TargetUpn: '',
    TargetUserId: null,
    Created: '2026-01-01T00:00:00Z',
    Modified: '2026-01-01T00:00:00Z'
  };
}

describe('SharePointDataService', () => {
  it('accepts a pre-built SPFI (web property detected via duck-typing)', () => {
    const sp = makeMockSp([], {});
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    expect(svc).toBeDefined();
  });

  it('getJobs selects the expected columns and orders by Id desc (always indexed)', async () => {
    const capture: ICapture = {};
    const sp = makeMockSp([], capture);
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    await svc.getJobs();
    expect(capture.select).toContain('Id');
    expect(capture.select).toContain('PayloadJson');
    expect(capture.select).toContain('JobType');
    expect(capture.orderBy).toEqual(['Id', false]);
  });

  it('getJobsPaged flags truncated when more rows exist beyond the page', async () => {
    const rows = Array.from({ length: 501 }, (_, i) => makeJobRow(i + 1));
    const sp = makeMockSp(rows, {});
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    const result = await svc.getJobsPaged(500);
    expect(result.truncated).toBe(true);
    expect(result.items).toHaveLength(500);

    const next = await result.next?.();
    expect(next?.items).toHaveLength(1);
    expect(next?.truncated).toBe(false);
  });

  it('getJobsPaged reports not truncated when fewer rows return', async () => {
    const rows = Array.from({ length: 10 }, (_, i) => makeJobRow(i + 1));
    const sp = makeMockSp(rows, {});
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    const result = await svc.getJobsPaged(500);
    expect(result.truncated).toBe(false);
    expect(result.items).toHaveLength(10);
  });

  it('getJobSummariesPaged excludes PayloadJson/StepsJson from the select', async () => {
    const capture: ICapture = {};
    const sp = makeMockSp([], capture);
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    await svc.getJobSummariesPaged();
    expect(capture.select).not.toContain('PayloadJson');
    expect(capture.select).not.toContain('StepsJson');
    expect(capture.select).toContain('TargetUpn');
  });

  it('getJobSummariesPaged applies a server-side filter for a search term', async () => {
    const capture: ICapture = {};
    const sp = makeMockSp([], capture);
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    await svc.getJobSummariesPaged({ search: 'anna' });
    expect(capture.filter).toContain('TargetUpn');
    expect(capture.filter).toContain('anna');
  });

  it('getTasksPaged follows the same continuation shape as getJobsPaged', async () => {
    const rows = Array.from({ length: 5 }, (_, i) => ({
      Id: i + 1,
      Title: `task-${i}`,
      JobId: null,
      TaskType: 'Other',
      Instructions: null,
      Status: 'Open',
      CompletedBy: null,
      CompletedUtc: null
    }));
    const sp = makeMockSp(rows, {});
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    const result = await svc.getTasksPaged(500);
    expect(result.truncated).toBe(false);
    expect(result.items).toHaveLength(5);
  });
});

describe('SharePointDataService column resilience', () => {
  const ALL_SUMMARY_COLUMNS: string[] = [
    'Id', 'Title', 'JobType', 'Status', 'CorrelationId', 'BatchId', 'TargetUpn',
    'ScheduledFor', 'Created', 'Modified', 'RunningSince', 'RequestedBy', 'ApprovedBy'
  ];

  it('selects every column when the list is fully up to date', async () => {
    const capture: ICapture = {};
    const sp = makeMockSp([makeJobRow(1)], capture, ALL_SUMMARY_COLUMNS);
    const svc = new SharePointDataService(sp as unknown as WebPartContext);

    await svc.getJobSummariesPaged();

    expect(capture.select).toContain('BatchId');
    expect(capture.select).toContain('RunningSince');
  });

  it('drops columns the list does not have yet instead of failing the whole query', async () => {
    const capture: ICapture = {};
    const outOfDate = ALL_SUMMARY_COLUMNS.filter((c) => c !== 'BatchId' && c !== 'RunningSince');
    const sp = makeMockSp([makeJobRow(1)], capture, outOfDate);
    const svc = new SharePointDataService(sp as unknown as WebPartContext);

    const page = await svc.getJobSummariesPaged();

    expect(capture.select).not.toContain('BatchId');
    expect(capture.select).not.toContain('RunningSince');
    expect(capture.select).toContain('Status');
    expect(page.items).toHaveLength(1);
  });

  it('defaults the dropped values so the summary still parses', async () => {
    const capture: ICapture = {};
    const outOfDate = ALL_SUMMARY_COLUMNS.filter((c) => c !== 'BatchId' && c !== 'RunningSince');
    const sp = makeMockSp([makeJobRow(7)], capture, outOfDate);
    const svc = new SharePointDataService(sp as unknown as WebPartContext);

    const page = await svc.getJobSummariesPaged();

    expect(page.items[0].batchId).toBe('');
    expect(page.items[0].runningSince).toBeNull();
    expect(page.items[0].itemId).toBe(7);
  });

  it('falls back to the full select when the field list cannot be read', async () => {
    const capture: ICapture = {};
    const sp = makeMockSp([makeJobRow(1)], capture);
    const svc = new SharePointDataService(sp as unknown as WebPartContext);

    await svc.getJobSummariesPaged();

    expect(capture.select).toContain('BatchId');
  });

  it('reads the field list once and caches it across queries', async () => {
    let fieldReads: number = 0;
    const capture: ICapture = {};
    const items = makeChain([makeJobRow(1)], capture);
    const fields = {
      select: (..._args: string[]) => ({
        top: (_n: number) => async () => {
          fieldReads++;
          return ALL_SUMMARY_COLUMNS.map((InternalName) => ({ InternalName }));
        }
      })
    };
    const sp = {
      web: {
        lists: { getByTitle: (_t: string) => ({ items, fields }) },
        currentUser: { select: (..._f: string[]) => async () => ({ Id: 1, Title: 'me', LoginName: 'me@contoso.com' }) }
      }
    };
    const svc = new SharePointDataService(sp as unknown as WebPartContext);

    await svc.getJobSummariesPaged();
    await svc.getJobSummariesPaged();

    expect(fieldReads).toBe(1);
  });
});
