
export type TelemetryLevel = 'info' | 'warning' | 'error';

const EMAIL_PATTERN: RegExp = /[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/gi;

function maskEmailLike(value: string): string {
  return value.replace(EMAIL_PATTERN, (match) => {
    const at: number = match.indexOf('@');
    return `${match[0]}***${match.slice(at)}`;
  });
}

export interface ITelemetryEvent {
  name: string;
  properties?: Record<string, string | number | boolean | undefined>;
  measurements?: Record<string, number>;
  level: TelemetryLevel;
  timestamp: string;
}

const RING_BUFFER_MAX: number = 200;

export class TelemetryService {
  private readonly _buffer: ITelemetryEvent[] = [];

  public trackEvent(name: string, properties?: Record<string, unknown>, level: TelemetryLevel = 'info'): void {
    const evt: ITelemetryEvent = {
      name,
      level,
      timestamp: new Date().toISOString(),
      properties: this._stringify(properties),
      measurements: this._measurements(properties)
    };
    this._buffer.push(evt);
    if (this._buffer.length > RING_BUFFER_MAX) {
      this._buffer.shift();
    }
    this._emit(evt);
  }

  public trackError(error: unknown, properties?: Record<string, unknown>): void {
    const err: Error = error instanceof Error ? error : new Error(String(error));
    this.trackEvent(
      'error',
      { ...(properties ?? {}), message: err.message, name: err.name },
      'error'
    );
  }

  public trackGraphCall(
    method: string,
    path: string,
    status: number,
    durationMs: number,
    retried: boolean
  ): void {
    this.trackEvent(
      'graph.call',
      { method, path: maskEmailLike(path).slice(0, 200), status, retried: String(retried) },
      status >= 400 ? 'warning' : 'info'
    );
  }

  public get events(): readonly ITelemetryEvent[] {
    return this._buffer;
  }

  private _emit(evt: ITelemetryEvent): void {
    const fn: typeof console.log =
      evt.level === 'error' ? console.error : evt.level === 'warning' ? console.warn : console.info;
    fn.call(console, `[UPC] ${evt.name}`, evt.properties ?? {});
  }

  private _stringify(props?: Record<string, unknown>): Record<string, string | number | boolean | undefined> | undefined {
    if (!props) {
      return undefined;
    }
    const out: Record<string, string | number | boolean | undefined> = {};
    for (const key of Object.keys(props)) {
      const v: unknown = props[key];
      if (v === undefined || v === null) {
        out[key] = undefined;
      } else if (typeof v === 'string') {
        out[key] = maskEmailLike(v);
      } else if (typeof v === 'number' || typeof v === 'boolean') {
        out[key] = v;
      } else {
        out[key] = maskEmailLike(String(v));
      }
    }
    return out;
  }

  private _measurements(props?: Record<string, unknown>): Record<string, number> | undefined {
    if (!props) {
      return undefined;
    }
    const out: Record<string, number> = {};
    for (const key of Object.keys(props)) {
      if (typeof props[key] === 'number') {
        out[key] = props[key] as number;
      }
    }
    return Object.keys(out).length > 0 ? out : undefined;
  }
}