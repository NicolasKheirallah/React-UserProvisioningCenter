export type AuditResult = 'Success' | 'Failure' | 'Skipped';

/** One UPC_AuditLog row. Every Graph write produces exactly one entry. */
export interface IAuditEntry {
  entryId: string;
  jobId: string;
  /** Operator UPN. */
  actor: string;
  action: string;
  targetUser: string;
  graphEndpoint: string;
  /** Secrets redacted before writing. */
  requestSummary: string;
  responseCode: number;
  durationMs: number;
  result: AuditResult;
  correlationId: string;
  timestampUtc: string;
}

/** What callers supply; ids and timestamps are added by AuditService. */
export type IAuditInput = Omit<IAuditEntry, 'entryId' | 'timestampUtc' | 'actor'>;
