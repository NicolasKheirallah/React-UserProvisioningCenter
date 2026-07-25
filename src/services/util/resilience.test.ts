jest.mock('./delay', () => ({
  delay: async (_ms: number, _signal?: AbortSignal): Promise<void> => undefined
}));

import { sharePointRetry, isRetryableSharePointError, getRetryAfterMs } from './sharePointRetry';
import { mapLimit } from './boundedConcurrency';
import { fetchPaged } from './pagedQuery';
import { ConcurrencyError } from './ConcurrencyError';

describe('sharePointRetry', () => {
  it('returns the result on the first attempt when the action succeeds', async () => {
    const result = await sharePointRetry(async () => 'ok');
    expect(result).toBe('ok');
  });

  it('retries transient SharePoint failures and eventually succeeds', async () => {
    let attempts = 0;
    const result = await sharePointRetry(
      async () => {
        attempts++;
        if (attempts < 3) {
          throw Object.assign(new Error('throttled'), { status: 429 });
        }
        return 'ok';
      },
      { maxAttempts: 4 }
    );
    expect(result).toBe('ok');
    expect(attempts).toBe(3);
  });

  it('gives up after maxAttempts and throws the last error', async () => {
    let attempts = 0;
    await expect(
      sharePointRetry(
        async () => {
          attempts++;
          throw Object.assign(new Error('unavailable'), { status: 503 });
        },
        { maxAttempts: 2 }
      )
    ).rejects.toThrow('unavailable');
    expect(attempts).toBe(2);
  });

  it('does not retry non-transient errors', async () => {
    let attempts = 0;
    await expect(
      sharePointRetry(async () => {
        attempts++;
        throw Object.assign(new Error('not found'), { status: 404 });
      })
    ).rejects.toThrow('not found');
    expect(attempts).toBe(1);
  });

  it('honors Retry-After when SharePoint sends it', async () => {
    let attempts = 0;
    const error = Object.assign(new Error('throttled'), {
      status: 429,
      response: {
        headers: {
          get: (name: string) => (name === 'Retry-After' ? '2' : null)
        }
      }
    });
    const result = await sharePointRetry(
      async () => {
        attempts++;
        if (attempts < 2) throw error;
        return 'ok';
      },
      { maxAttempts: 3 }
    );
    expect(result).toBe('ok');
    expect(attempts).toBe(2);
  });

  it('aborts when the signal is already aborted', async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(
      sharePointRetry(async () => 'ok', { signal: controller.signal })
    ).rejects.toThrow('SharePoint request aborted');
  });
});

describe('isRetryableSharePointError', () => {
  it('treats 429, 503, 504 and status 0 as retryable', () => {
    expect(isRetryableSharePointError({ status: 429 })).toBe(true);
    expect(isRetryableSharePointError({ status: 503 })).toBe(true);
    expect(isRetryableSharePointError({ status: 504 })).toBe(true);
    expect(isRetryableSharePointError({ status: 0 })).toBe(true);
  });

  it('treats 404 as not retryable', () => {
    expect(isRetryableSharePointError({ status: 404 })).toBe(false);
  });

  it('treats abort errors as not retryable', () => {
    expect(isRetryableSharePointError(new DOMException('aborted', 'AbortError'))).toBe(false);
  });
});

describe('getRetryAfterMs', () => {
  it('returns milliseconds from Retry-After header', () => {
    const err = {
      response: {
        headers: {
          get: (name: string) => (name === 'Retry-After' ? '5' : null)
        }
      }
    };
    expect(getRetryAfterMs(err)).toBe(5000);
  });

  it('returns undefined when the header is missing', () => {
    expect(getRetryAfterMs({ response: { headers: { get: () => null } } })).toBeUndefined();
  });
});

describe('mapLimit', () => {
  it('respects the concurrency limit', async () => {
    let running = 0;
    let maxRunning = 0;
    const results = await mapLimit([1, 2, 3, 4, 5], 2, async (n) => {
      running++;
      maxRunning = Math.max(maxRunning, running);
      await new Promise((resolve) => setTimeout(resolve, 10));
      running--;
      return n * 2;
    });
    expect(results).toEqual([2, 4, 6, 8, 10]);
    expect(maxRunning).toBe(2);
  });

  it('preserves order in the output array', async () => {
    const results = await mapLimit([3, 2, 1], 2, async (n) => {
      await new Promise((resolve) => setTimeout(resolve, 20 - n * 5));
      return n;
    });
    expect(results).toEqual([3, 2, 1]);
  });
});

describe('fetchPaged', () => {
  function makePaged<T>(rows: T[], pageSize: number): AsyncIterable<T[]> {
    let remaining = [...rows];
    return {
      [Symbol.asyncIterator](): AsyncIterator<T[]> {
        return {
          next: async () => {
            if (remaining.length === 0) {
              return { done: true, value: undefined };
            }
            const page = remaining.slice(0, pageSize);
            remaining = remaining.slice(page.length);
            return { done: false, value: page };
          }
        };
      }
    };
  }

  it('returns the first page and a continuation callback', async () => {
    const result = await fetchPaged(makePaged([1, 2, 3], 2), 2);
    expect(result.items).toEqual([1, 2]);
    expect(result.truncated).toBe(true);
    expect(result.next).toBeDefined();
  });

  it('returns not truncated when rows are fewer than pageSize', async () => {
    const result = await fetchPaged(makePaged([1, 2], 3), 3);
    expect(result.items).toEqual([1, 2]);
    expect(result.truncated).toBe(false);
    expect(result.next).toBeUndefined();
  });

  it('uses the retry wrapper for the initial fetch and continuation', async () => {
    const iterable = makePaged([1, 2], 1);
    let retryCalls = 0;
    const retry = async <R>(action: () => Promise<R>): Promise<R> => {
      retryCalls++;
      return action();
    };
    const result = await fetchPaged(iterable, 1, retry);
    expect(result.items).toEqual([1]);
    expect(result.truncated).toBe(true);
    const nextResult = await result.next?.();
    expect(nextResult?.items).toEqual([2]);
    expect(retryCalls).toBe(2);
  });
});

describe('ConcurrencyError', () => {
  it('carries the provided message and name', () => {
    const err = new ConcurrencyError('job is locked');
    expect(err.message).toBe('job is locked');
    expect(err.name).toBe('ConcurrencyError');
  });
});
