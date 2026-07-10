jest.mock('@pnp/sp', () => ({ spfi: () => ({ using: () => ({}) }), SPFx: () => ({}), __esModule: true }));
jest.mock('@pnp/sp/webs', () => ({ __esModule: true }));
jest.mock('@pnp/sp/lists', () => ({ __esModule: true }));
jest.mock('@pnp/sp/items', () => ({ __esModule: true }));
jest.mock('@pnp/sp/site-users/web', () => ({ __esModule: true }));

import { SharePointDataService } from './SharePointDataService';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

interface ICapture {
  select?: string[];
  expand?: string[];
  filter?: string;
  orderBy?: [string, boolean];
  top?: number;
}

/** Build a PnPJS-style chainable + callable object. Every method returns self. */
function makeChain(returnValue: unknown[], capture: ICapture): unknown {
  const fn = (async (): Promise<unknown[]> => returnValue) as unknown as (
    ..._args: unknown[]
  ) => Promise<unknown[]>;
  const chain = fn as unknown as Record<string, unknown>;
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
  chain.getById = (_id: number): unknown => {
    return {
      ...chain,
      update: async (_payload: unknown): Promise<void> => undefined
    };
  };
  return chain;
}

function makeMockSp(returnValue: unknown[], capture: ICapture): unknown {
  const items = makeChain(returnValue, capture);
  const lists = {
    getByTitle: (_title: string) => ({ items }),
    ensure: async () => ({ created: true, list: {} })
  };
  const web = {
    lists,
    currentUser: {
      select: (_field: string) => (async () => ({ Id: 1, Title: 'me' })) as () => Promise<unknown>
    }
  };
  return { web, __root: 'mock' };
}

describe('SharePointDataService', () => {
  it('accepts a pre-built SPFI (web property detected via duck-typing)', () => {
    const sp = makeMockSp([], {});
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    expect(svc).toBeDefined();
  });

  it('getJobs selects the expected columns and orders by Modified desc', async () => {
    const capture: ICapture = {};
    const sp = makeMockSp([], capture);
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    await svc.getJobs();
    expect(capture.select).toContain('Id');
    expect(capture.select).toContain('PayloadJson');
    expect(capture.select).toContain('JobType');
    expect(capture.orderBy).toEqual(['Modified', false]);
  });

  it('getJobsPaged flags truncated when the result equals the page size', async () => {
    // getJobsPaged asks for top+1 (501) to detect truncation without a false
    // positive at exactly `top` rows; the mock ignores .top() and returns
    // whatever it's given, so 501 rows here simulates the server returning
    // one more than requested.
    const rows = Array.from({ length: 501 }, (_, i) => ({
      Id: i + 1,
      Title: `job-${i}`,
      JobType: 'Onboard',
      Status: 'Completed',
      PayloadJson: '{}',
      StepsJson: '[]',
      ScheduledFor: null,
      CorrelationId: 'c',
      TargetUserId: null,
      Created: '2026-01-01T00:00:00Z'
    }));
    const sp = makeMockSp(rows, {});
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    const result = await svc.getJobsPaged(500);
    expect(result.truncated).toBe(true);
    expect(result.items).toHaveLength(500);
  });

  it('getJobsPaged reports not truncated when fewer rows return', async () => {
    const rows = Array.from({ length: 10 }, (_, i) => ({
      Id: i + 1,
      Title: `job-${i}`,
      JobType: 'Onboard',
      Status: 'Completed',
      PayloadJson: '{}',
      StepsJson: '[]',
      ScheduledFor: null,
      CorrelationId: 'c',
      TargetUserId: null,
      Created: '2026-01-01T00:00:00Z'
    }));
    const sp = makeMockSp(rows, {});
    const svc = new SharePointDataService(sp as unknown as WebPartContext);
    const result = await svc.getJobsPaged(500);
    expect(result.truncated).toBe(false);
    expect(result.items).toHaveLength(10);
  });
});