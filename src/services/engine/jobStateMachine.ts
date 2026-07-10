import type { JobStatus } from '../../models';

/**
 * Job status state machine (spec Section 4 statuses). Client-enforced:
 * with delegated permissions this guards workflow correctness, not security
 * (spec Section 1 — approval is advisory by design).
 */
const TRANSITIONS: Record<JobStatus, JobStatus[]> = {
  Draft: ['PendingApproval', 'Cancelled'],
  PendingApproval: ['Approved', 'Cancelled'],
  Approved: ['Scheduled', 'Running', 'Cancelled'],
  Scheduled: ['Running', 'Cancelled'],
  Running: ['Completed', 'PartiallyFailed', 'Failed', 'Cancelled'],
  PartiallyFailed: ['Running', 'Completed', 'Cancelled'],
  Failed: ['Running', 'Cancelled'],
  Completed: [],
  Cancelled: []
};

/** Statuses from which the engine may begin (or resume) executing steps. */
const STARTABLE: JobStatus[] = ['Approved', 'Scheduled', 'Running', 'PartiallyFailed', 'Failed'];

export function canTransition(from: JobStatus, to: JobStatus): boolean {
  return (TRANSITIONS[from] ?? []).indexOf(to) !== -1;
}

export function assertTransition(from: JobStatus, to: JobStatus): void {
  if (!canTransition(from, to)) {
    throw new Error(`Invalid job transition ${from} -> ${to}`);
  }
}

/** The engine refuses to start anything not yet approved (Phase 1 acceptance). */
export function canStartJob(status: JobStatus): boolean {
  return STARTABLE.indexOf(status) !== -1;
}

export function isTerminal(status: JobStatus): boolean {
  return TRANSITIONS[status].length === 0;
}
