import { TelemetryService } from './TelemetryService';

describe('TelemetryService', () => {
  let service: TelemetryService;
  let consoleSpy: jest.SpyInstance;

  beforeEach(() => {
    service = new TelemetryService();
    consoleSpy = jest.spyOn(console, 'info').mockImplementation(() => undefined);
  });

  afterEach(() => {
    consoleSpy.mockRestore();
  });

  it('captures events in the ring buffer', () => {
    service.trackEvent('test.event', { foo: 'bar' });
    service.trackEvent('test.event2', { count: 5 });
    expect(service.events).toHaveLength(2);
    expect(service.events[0].name).toBe('test.event');
    expect(service.events[1].name).toBe('test.event2');
  });

  it('emits to console at the right level', () => {
    const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
    service.trackEvent('test.warn', {}, 'warning');
    expect(warnSpy).toHaveBeenCalled();
    warnSpy.mockRestore();
  });

  it('tracks errors with the error name and message', () => {
    const errSpy = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    service.trackError(new Error('boom'), { context: 'ctx' });
    expect(service.events).toHaveLength(1);
    expect(service.events[0].name).toBe('error');
    expect(service.events[0].level).toBe('error');
    expect(service.events[0].properties?.message).toBe('boom');
    errSpy.mockRestore();
  });

  it('tracks Graph calls with status and duration', () => {
    service.trackGraphCall('GET', '/users', 200, 42, false);
    expect(service.events).toHaveLength(1);
    expect(service.events[0].name).toBe('graph.call');
    expect(service.events[0].properties?.method).toBe('GET');
    expect(service.events[0].properties?.status).toBe(200);
    expect(service.events[0].level).toBe('info');
  });

  it('elevates level to warning for failed Graph calls', () => {
    service.trackGraphCall('POST', '/users', 500, 100, true);
    expect(service.events[0].level).toBe('warning');
  });

  it('rotates the ring buffer at 200 entries', () => {
    for (let i = 0; i < 210; i++) {
      service.trackEvent(`event.${i}`);
    }
    expect(service.events).toHaveLength(200);
    expect(service.events[0].name).toBe('event.10');
    expect(service.events[199].name).toBe('event.209');
  });

  it('stringifies non-primitive property values', () => {
    service.trackEvent('test', { obj: { nested: true } });
    expect(service.events[0].properties?.obj).toBe('[object Object]');
  });

  it('extracts numeric properties into measurements', () => {
    service.trackEvent('test', { durationMs: 500, count: 3 });
    expect(service.events[0].measurements?.durationMs).toBe(500);
    expect(service.events[0].measurements?.count).toBe(3);
  });
});