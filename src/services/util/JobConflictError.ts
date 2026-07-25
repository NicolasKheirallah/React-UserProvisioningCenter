/**
 * Raised when a SharePoint write is rejected with HTTP 412 because the item's
 * ETag moved — another tab, another operator, or a background flow changed the
 * job between our read and our write.
 *
 * Previously a 412 surfaced as an opaque PnPJS error with a message no
 * operator could act on. Modelling it explicitly lets the UI say "this job
 * changed elsewhere, refresh to see the latest state" and lets callers choose
 * to re-read and re-apply rather than fail the whole run.
 */
export class JobConflictError extends Error {
  public readonly itemId: number;

  public constructor(itemId: number, message?: string) {
    super(message ?? `Job ${itemId} was modified in another session`);
    Object.setPrototypeOf(this, JobConflictError.prototype);
    this.name = 'JobConflictError';
    this.itemId = itemId;
  }
}

/** True for the PnPJS/SharePoint shape of an ETag precondition failure. */
export function isEtagConflict(err: unknown): boolean {
  if (err instanceof JobConflictError) {
    return true;
  }
  return (
    typeof err === 'object' &&
    err !== null &&
    'status' in err &&
    (err as { status: unknown }).status === 412
  );
}
