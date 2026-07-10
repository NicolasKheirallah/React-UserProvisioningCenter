export const RETRYABLE_STATUS_CODES: readonly number[] = [429, 502, 503, 504];

/**
 * Normalized Graph failure. `requestId` is the Graph request-id used as the
 * correlation id on error surfaces — raw payloads are never shown to users.
 */
export class GraphServiceError extends Error {
  public readonly statusCode: number;
  public readonly graphCode: string;
  public readonly requestId: string;
  public readonly retryable: boolean;
  /** Seconds to wait, when the service sent Retry-After. */
  public readonly retryAfterSeconds: number | undefined;

  public constructor(
    message: string,
    statusCode: number,
    graphCode: string,
    requestId: string,
    retryAfterSeconds?: number
  ) {
    super(message);
    // Restore the prototype chain — required for instanceof to work when
    // classes extending Error are downleveled by the build target.
    Object.setPrototypeOf(this, GraphServiceError.prototype);
    this.name = 'GraphServiceError';
    this.statusCode = statusCode;
    this.graphCode = graphCode;
    this.requestId = requestId;
    // statusCode 0 means the SDK threw a raw network/offline/CORS error
    // (no HTTP response was ever received) rather than a real Graph
    // rejection — treat it as transient and retryable, same as 429/5xx.
    this.retryable = statusCode === 0 || RETRYABLE_STATUS_CODES.indexOf(statusCode) !== -1;
    this.retryAfterSeconds = retryAfterSeconds;
  }
}

/** Thrown when a caller aborts via AbortSignal (component unmount). */
export class RequestAbortedError extends Error {
  public constructor() {
    super('Request aborted');
    Object.setPrototypeOf(this, RequestAbortedError.prototype);
    this.name = 'RequestAbortedError';
  }
}

interface IRawGraphErrorShape {
  statusCode?: number;
  code?: string;
  requestId?: string;
  message?: string;
  body?: string;
  headers?: { get?: (name: string) => string | null };
}

/** Convert whatever the Graph SDK threw into a GraphServiceError. */
export function toGraphServiceError(err: unknown): GraphServiceError {
  if (err instanceof GraphServiceError) {
    return err;
  }
  const raw: IRawGraphErrorShape = (err ?? {}) as IRawGraphErrorShape;
  let code: string = raw.code ?? 'UnknownError';
  let message: string = raw.message ?? 'Unknown Graph error';
  // The SDK sometimes puts the OData error JSON in body.
  if (raw.body && typeof raw.body === 'string') {
    try {
      const parsed: { error?: { code?: string; message?: string } } = JSON.parse(raw.body);
      code = parsed.error?.code ?? code;
      message = parsed.error?.message ?? message;
    } catch {
      // body was not JSON; keep defaults
    }
  }
  let retryAfter: number | undefined;
  const headerValue: string | null | undefined = raw.headers?.get?.('Retry-After');
  if (headerValue) {
    const parsedHeader: number = parseInt(headerValue, 10);
    if (!isNaN(parsedHeader)) {
      retryAfter = parsedHeader;
    }
  }
  return new GraphServiceError(
    message,
    typeof raw.statusCode === 'number' ? raw.statusCode : 0,
    code,
    raw.requestId ?? '',
    retryAfter
  );
}
