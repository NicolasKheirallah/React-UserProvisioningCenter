import type { JobStatus, JobType } from './IJob';
export interface IJobQuery {
  search?: string;
  status?: JobStatus[];
  jobType?: JobType[];
  createdFromUtc?: string;
  createdToUtc?: string;
  requestedBy?: string;
  top?: number;
}
export interface IJobSummary {
  itemId: number;
  jobId: string;
  jobType: JobType;
  status: JobStatus;
  targetUpn: string;
  scheduledFor: string | null;
  requestedBy: string | null;
  approvedBy: string | null;
  correlationId: string;
  createdUtc: string | null;
  modifiedUtc: string | null;
}
export function isEmptyJobQuery(query: IJobQuery | undefined): boolean {
  if (!query) {
    return true;
  }
  return (
    !query.search?.trim() &&
    (query.status ?? []).length === 0 &&
    (query.jobType ?? []).length === 0 &&
    !query.createdFromUtc &&
    !query.createdToUtc &&
    !query.requestedBy
  );
}
