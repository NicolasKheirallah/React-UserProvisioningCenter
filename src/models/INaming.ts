export type CandidateRejectionReason =
  | 'reserved'
  | 'collision'
  | 'collision-soft-deleted'
  | 'invalid';

export interface IRejectedCandidate {
  candidate: string;
  reason: CandidateRejectionReason;
}

/** Result of the naming policy engine (spec Section 5). */
export interface INamingResult {
  /** First available candidate, or null when everything up to the search budget collided. */
  chosen: string | null;
  rejected: IRejectedCandidate[];
  /** Up to 3 additional available local parts the operator may pick instead. */
  alternatives: string[];
}

/** Availability verdict for one UPN local part. */
export type UpnAvailability = 'available' | 'taken' | 'taken-soft-deleted';

/** Injected by the service; pure engine code never talks to Graph directly. */
export type AvailabilityChecker = (
  upn: string,
  signal?: AbortSignal
) => Promise<UpnAvailability>;
