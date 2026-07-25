import { shouldRetryQuery } from './queryRetry';
import { TimeoutError } from './withTimeout';
import { ConcurrencyError } from './ConcurrencyError';
import { CircuitOpenError } from './circuitBreaker';

describe('shouldRetryQuery', () => {
  it('never retries a timeout, so a hung request surfaces once rather than doubling the wait', () => {
    expect(shouldRetryQuery(0, new TimeoutError(20_000, 'UPC_ProvisioningJobs'))).toBe(false);
  });

  it('never retries an open circuit', () => {
    expect(shouldRetryQuery(0, new CircuitOpenError('UPC_ProvisioningJobs', 5000))).toBe(false);
  });

  it('never retries a concurrency conflict', () => {
    expect(shouldRetryQuery(0, new ConcurrencyError('locked'))).toBe(false);
  });

  it('does not retry client errors such as a missing list or column', () => {
    expect(shouldRetryQuery(0, { status: 404 })).toBe(false);
    expect(shouldRetryQuery(0, { status: 400 })).toBe(false);
    expect(shouldRetryQuery(0, { status: 403 })).toBe(false);
  });

  it('still retries throttling and request timeout', () => {
    expect(shouldRetryQuery(0, { status: 429 })).toBe(true);
    expect(shouldRetryQuery(0, { status: 408 })).toBe(true);
  });

  it('retries server errors once', () => {
    expect(shouldRetryQuery(0, { status: 503 })).toBe(true);
    expect(shouldRetryQuery(1, { status: 503 })).toBe(false);
  });

  it('reads statusCode as well as status', () => {
    expect(shouldRetryQuery(0, { statusCode: 404 })).toBe(false);
  });

  it('retries an unrecognised error once', () => {
    expect(shouldRetryQuery(0, new Error('network glitch'))).toBe(true);
    expect(shouldRetryQuery(1, new Error('network glitch'))).toBe(false);
  });
});
