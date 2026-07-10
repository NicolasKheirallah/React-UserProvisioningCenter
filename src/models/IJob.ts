import type { IOnboardingPayload } from './IOnboardingPayload';
import type { IOffboardingPayload } from './IOffboardingPayload';
import type { ITransferPayload } from './ITransferPayload';

/**
 * Clone reuses the onboarding pipeline (new hire, same shape) plus a
 * cloneSourceUserId to copy an existing user's licenses/groups from —
 * routed by IProvisioningJob.jobType, not a distinct payload kind. Bulk rows
 * are submitted as individual Onboard jobs, one per row.
 */
export type JobType = 'Onboard' | 'Offboard' | 'Transfer' | 'Clone' | 'Bulk';

/**
 * Any job payload. Each variant carries an explicit `kind` discriminant so a
 * new payload type can't silently mis-narrow by falling through a "not the
 * other kind" guard.
 */
export type IJobPayload = IOnboardingPayload | IOffboardingPayload | ITransferPayload;

export function isOffboardingPayload(payload: IJobPayload): payload is IOffboardingPayload {
  return payload.kind === 'offboard';
}

export function isOnboardingPayload(payload: IJobPayload): payload is IOnboardingPayload {
  return payload.kind === 'onboard';
}

export function isTransferPayload(payload: IJobPayload): payload is ITransferPayload {
  return payload.kind === 'transfer';
}

export type JobStatus =
  | 'Draft'
  | 'PendingApproval'
  | 'Approved'
  | 'Scheduled'
  | 'Running'
  | 'PartiallyFailed'
  | 'Failed'
  | 'Completed'
  | 'Cancelled';

export type StepStatus = 'pending' | 'running' | 'completed' | 'failed' | 'skipped';

export interface IStepError {
  graphCode: string;
  message: string;
  retryable: boolean;
}

/** One element of StepsJson (Section 4 of the spec). */
export interface IJobStep {
  stepId: string;
  status: StepStatus;
  attempts: number;
  maxAttempts: number;
  lastError: IStepError | null;
  startedUtc: string | null;
  completedUtc: string | null;
  skippable: boolean;
}

/** Parsed representation of a UPC_ProvisioningJobs item. */
export interface IProvisioningJob {
  /** SharePoint list item id. */
  itemId: number;
  /** Title column — job GUID. */
  jobId: string;
  jobType: JobType;
  status: JobStatus;
  /** Parsed PayloadJson. Immutable after approval; never contains secrets. */
  payload: IJobPayload;
  /** Parsed StepsJson. */
  steps: IJobStep[];
  /** ISO string; null/undefined = run immediately. */
  scheduledFor: string | null;
  requestedBy: string | null;
  approvedBy: string | null;
  correlationId: string;
  /** Entra objectId once CreateUser has run. */
  targetUserId: string | null;
  /** ISO timestamp of the SharePoint item's Created field. */
  createdUtc: string | null;
}
