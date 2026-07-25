import { QueryClient } from '@tanstack/react-query';
import { shouldRetryQuery } from '../services/util/queryRetry';

function makeClient(): QueryClient {
  return new QueryClient({
    defaultOptions: {
      queries: { retry: shouldRetryQuery, refetchOnWindowFocus: false, networkMode: 'always' },
      mutations: { networkMode: 'always' }
    }
  });
}

describe('QueryClient network mode', () => {
  it("uses networkMode 'always' so queries are not paused when navigator.onLine is false", () => {
    const defaults = makeClient().getDefaultOptions();
    expect(defaults.queries?.networkMode).toBe('always');
  });

  it("applies networkMode 'always' to mutations too", () => {
    const defaults = makeClient().getDefaultOptions();
    expect(defaults.mutations?.networkMode).toBe('always');
  });

  it('does not leave networkMode at the v4 default, which pauses every query offline', () => {
    const defaults = makeClient().getDefaultOptions();
    expect(defaults.queries?.networkMode).not.toBe('online');
    expect(defaults.queries?.networkMode).not.toBeUndefined();
  });
});
