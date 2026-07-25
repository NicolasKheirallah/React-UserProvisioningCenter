/**
 * Thrown when an engine operation cannot proceed because another session
 * currently holds the job lock, or because an optimistic-concurrency/ETag
 * check failed.
 */
export class ConcurrencyError extends Error {
  public constructor(message: string) {
    super(message);
    Object.setPrototypeOf(this, ConcurrencyError.prototype);
    this.name = 'ConcurrencyError';
  }
}
