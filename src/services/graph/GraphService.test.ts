import { GraphService, batchTyped, type IBatchResponse } from './GraphService';
import { GraphServiceError, RequestAbortedError } from './GraphError';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { TelemetryService } from '../telemetry/TelemetryService';

type Responder = () => unknown;

/**
 * Mocks the MSGraphClientV3 fluent API (`.api(path).version(v).headers(h).get()`),
 * not GraphService itself — so GraphService's own retry/backoff/timeout/batch
 * code in _execute/_send/batch actually runs under test, the way
 * WorkflowEngine.test.ts's MockGraph (which stubs GraphService's public
 * methods directly) never exercises.
 */
class MockGraphClient {
  public calls: { method: string; path: string; body?: unknown }[] = [];
  private readonly _queues: Map<string, Responder[]> = new Map();

  /** Queue the next response (or thrown error) for one method+path. FIFO per key. */
  public queue(method: string, path: string, responder: Responder): void {
    const key = `${method} ${path}`;
    const list = this._queues.get(key) ?? [];
    list.push(responder);
    this._queues.set(key, list);
  }

  public api(path: string): unknown {
    const builder = {
      version: (): unknown => builder,
      headers: (): unknown => builder,
      get: (): Promise<unknown> => this._dispatch('GET', path),
      post: (body: unknown): Promise<unknown> => this._dispatch('POST', path, body),
      patch: (body: unknown): Promise<unknown> => this._dispatch('PATCH', path, body),
      put: (body: unknown): Promise<unknown> => this._dispatch('PUT', path, body),
      delete: (): Promise<unknown> => this._dispatch('DELETE', path)
    };
    return builder;
  }

  private async _dispatch(method: string, path: string, body?: unknown): Promise<unknown> {
    this.calls.push({ method, path, body });
    const key = `${method} ${path}`;
    const queue = this._queues.get(key);
    const responder = queue?.shift();
    if (!responder) {
      throw new Error(`MockGraphClient: no queued responder left for ${key}`);
    }
    return responder();
  }

  public callCount(method: string, path: string): number {
    return this.calls.filter((c) => c.method === method && c.path === path).length;
  }
}

function makeGraph(): { graph: GraphService; client: MockGraphClient } {
  const client = new MockGraphClient();
  const graph = new GraphService(client as unknown as MSGraphClientV3);
  return { graph, client };
}

/** Shape the SDK throws for a Graph-rejected request. */
function graphError(statusCode: number, code: string, retryAfterSeconds?: string): unknown {
  return {
    statusCode,
    code,
    message: code,
    requestId: 'req-1',
    headers: {
      get: (name: string): string | null =>
        name === 'Retry-After' && retryAfterSeconds !== undefined ? retryAfterSeconds : null
    }
  };
}

/** Shape a raw network/offline/CORS failure — no statusCode, no headers. */
function networkError(): unknown {
  return { message: 'Failed to fetch' };
}

describe('GraphService._execute (retry/backoff)', () => {
  it('returns the body on a first-attempt success with exactly one call', async () => {
    const { graph, client } = makeGraph();
    client.queue('GET', '/me', () => ({ id: 'me-1' }));

    const result = await graph.get<{ id: string }>('/me');

    expect(result).toEqual({ id: 'me-1' });
    expect(client.callCount('GET', '/me')).toBe(1);
  });

  it('retries a 429 honoring Retry-After, then succeeds', async () => {
    const { graph, client } = makeGraph();
    client.queue('GET', '/users', () => {
      throw graphError(429, 'TooManyRequests', '0');
    });
    client.queue('GET', '/users', () => ({ value: [] }));

    const result = await graph.get<{ value: unknown[] }>('/users');

    expect(result).toEqual({ value: [] });
    expect(client.callCount('GET', '/users')).toBe(2);
  });

  it('does not retry a non-retryable 4xx — fails after exactly one attempt', async () => {
    const { graph, client } = makeGraph();
    client.queue('GET', '/users/x', () => {
      throw graphError(400, 'BadRequest');
    });

    await expect(graph.get('/users/x')).rejects.toMatchObject({ statusCode: 400, retryable: false });
    expect(client.callCount('GET', '/users/x')).toBe(1);
  });

  it('treats a raw network failure (no statusCode) as retryable', async () => {
    const { graph, client } = makeGraph();
    client.queue('GET', '/organization', () => {
      throw networkError();
    });
    client.queue('GET', '/organization', () => ({ id: 'org-1' }));

    const result = await graph.get<{ id: string }>('/organization');

    expect(result).toEqual({ id: 'org-1' });
    expect(client.callCount('GET', '/organization')).toBe(2);
  }, 10000);

  it('gives up after maxAttempts and throws the last error', async () => {
    const { graph, client } = makeGraph();
    for (let i = 0; i < 3; i++) {
      client.queue('GET', '/flaky', () => {
        throw graphError(503, 'ServiceUnavailable', '0');
      });
    }

    await expect(graph.get('/flaky', { maxAttempts: 3 })).rejects.toMatchObject({
      statusCode: 503,
      retryable: true
    });
    expect(client.callCount('GET', '/flaky')).toBe(3);
  });

  it('rejects immediately without calling the client when the signal is already aborted', async () => {
    const { graph, client } = makeGraph();
    const controller = new AbortController();
    controller.abort();

    await expect(graph.get('/me', { signal: controller.signal })).rejects.toBeInstanceOf(
      RequestAbortedError
    );
    expect(client.calls).toHaveLength(0);
  });

  it('propagates abort requested during the retry backoff wait', async () => {
    const { graph, client } = makeGraph();
    const controller = new AbortController();
    client.queue('GET', '/slow-retry', () => {
      throw graphError(503, 'ServiceUnavailable'); // no Retry-After -> exponential backoff
    });

    const promise = graph.get('/slow-retry', { signal: controller.signal });
    // Abort while the first failure's backoff delay is in flight.
    setTimeout(() => controller.abort(), 5);

    await expect(promise).rejects.toBeInstanceOf(RequestAbortedError);
    expect(client.callCount('GET', '/slow-retry')).toBe(1);
  });

  it('reports every attempt to telemetry, flagging retried attempts', async () => {
    const { graph, client } = makeGraph();
    client.queue('GET', '/me', () => {
      throw graphError(429, 'TooManyRequests', '0');
    });
    client.queue('GET', '/me', () => ({ id: 'me-1' }));
    const trackGraphCall = jest.fn();
    graph.setTelemetry({ trackGraphCall } as unknown as TelemetryService);

    await graph.get('/me');

    expect(trackGraphCall).toHaveBeenCalledTimes(2);
    expect(trackGraphCall).toHaveBeenNthCalledWith(1, 'GET', '/me', 429, expect.any(Number), false);
    expect(trackGraphCall).toHaveBeenNthCalledWith(2, 'GET', '/me', 200, expect.any(Number), true);
  });
});

describe('GraphService._send (per-request timeout)', () => {
  it('rejects with a 408 GraphServiceError when the call outlasts timeoutMs', async () => {
    const { graph, client } = makeGraph();
    client.queue('GET', '/hanging', () => new Promise(() => undefined)); // never resolves

    await expect(graph.get('/hanging', { timeoutMs: 20, maxAttempts: 1 })).rejects.toMatchObject({
      statusCode: 408,
      graphCode: 'RequestTimeout'
    });
  });

  it('resolves normally when the call finishes before timeoutMs', async () => {
    const { graph, client } = makeGraph();
    client.queue(
      'GET',
      '/fast',
      () => new Promise((resolve) => setTimeout(() => resolve({ id: 'ok' }), 5))
    );

    const result = await graph.get<{ id: string }>('/fast', { timeoutMs: 200 });
    expect(result).toEqual({ id: 'ok' });
  });
});

describe('GraphService.batch (sub-response retry)', () => {
  it('does not retry when every sub-response already succeeded', async () => {
    const { graph, client } = makeGraph();
    client.queue('POST', '/$batch', () => ({
      responses: [
        { id: 'a', status: 200, body: { ok: true } },
        { id: 'b', status: 200, body: { ok: true } }
      ]
    }));

    const result = await graph.batch([
      { id: 'a', method: 'GET', url: '/a' },
      { id: 'b', method: 'GET', url: '/b' }
    ]);

    expect(result.get('a')?.status).toBe(200);
    expect(result.get('b')?.status).toBe(200);
    expect(client.callCount('POST', '/$batch')).toBe(1);
  });

  it('retries only the throttled sub-request, honoring its own Retry-After, and merges the final result', async () => {
    const { graph, client } = makeGraph();
    client.queue('POST', '/$batch', (): { responses: IBatchResponse[] } => ({
      responses: [
        { id: 'a', status: 429, body: {}, headers: { 'Retry-After': '0' } },
        { id: 'b', status: 200, body: { ok: 'b' } }
      ]
    }));
    client.queue('POST', '/$batch', (): { responses: IBatchResponse[] } => ({
      responses: [{ id: 'a', status: 200, body: { ok: 'a' } }]
    }));

    const result = await graph.batch([
      { id: 'a', method: 'GET', url: '/a' },
      { id: 'b', method: 'GET', url: '/b' }
    ]);

    expect(result.get('a')).toEqual({ id: 'a', status: 200, body: { ok: 'a' } });
    expect(result.get('b')?.status).toBe(200);
    expect(client.callCount('POST', '/$batch')).toBe(2);
    // The second round only re-sent the throttled sub-request, not 'b' again.
    const secondRoundBody = client.calls[1].body as { requests: { id: string }[] };
    expect(secondRoundBody.requests.map((r) => r.id)).toEqual(['a']);
  });

  it('stops after maxRounds even if a sub-request is still throttled, keeping the last-seen status', async () => {
    const { graph, client } = makeGraph();
    for (let round = 0; round < 3; round++) {
      client.queue('POST', '/$batch', (): { responses: IBatchResponse[] } => ({
        responses: [{ id: 'a', status: 429, body: {}, headers: { 'Retry-After': '0' } }]
      }));
    }

    const result = await graph.batch([{ id: 'a', method: 'GET', url: '/a' }]);

    expect(result.get('a')?.status).toBe(429);
    expect(client.callCount('POST', '/$batch')).toBe(3);
  });
});

describe('batchTyped', () => {
  it('maps typed request keys to ids and back to a typed response object', async () => {
    const { graph, client } = makeGraph();
    client.queue('POST', '/$batch', () => ({
      responses: [
        { id: 'org', status: 200, body: { id: 'org-1' } },
        { id: 'me', status: 200, body: { id: 'me-1' } }
      ]
    }));

    const result = await batchTyped(graph, {
      org: { method: 'GET', url: '/organization?$select=id' },
      me: { method: 'GET', url: '/me?$select=id' }
    });

    expect(result.org.body).toEqual({ id: 'org-1' });
    expect(result.me.body).toEqual({ id: 'me-1' });
    const sentIds = (client.calls[0].body as { requests: { id: string }[] }).requests.map((r) => r.id);
    expect(sentIds.sort()).toEqual(['me', 'org']);
  });

  it('falls back to a zero-status placeholder for a key Graph never answered', async () => {
    const { graph, client } = makeGraph();
    client.queue('POST', '/$batch', () => ({
      responses: [{ id: 'org', status: 200, body: { id: 'org-1' } }]
      // 'me' intentionally missing from the response.
    }));

    const result = await batchTyped(graph, {
      org: { method: 'GET', url: '/organization' },
      me: { method: 'GET', url: '/me' }
    });

    expect(result.me).toEqual({ id: 'me', status: 0, body: undefined });
  });
});

describe('GraphServiceError shape', () => {
  it('is thrown (not the raw SDK error) so callers get a normalized, typed error', async () => {
    const { graph, client } = makeGraph();
    client.queue('GET', '/boom', () => {
      throw graphError(403, 'Forbidden');
    });

    await expect(graph.get('/boom')).rejects.toBeInstanceOf(GraphServiceError);
  });
});
