/**
 * Approval quorum and delegation (RBAC enhancements).
 *
 * Both are *workflow* controls, not security boundaries. With delegated Graph
 * permissions, effective access is always the operator's own Entra role — a
 * quorum makes an approval policy explicit and auditable, it does not stop a
 * User Administrator from doing the same thing by hand in the portal. Say so
 * plainly in the admin docs rather than implying an enforcement guarantee.
 */
export interface IApprovalRecord {
  actor: string;
  actorUpn: string;
  timestampUtc: string;
  onBehalfOf?: string;
}
export interface IApprovalDelegation {
  itemId: number;
  delegatorUpn: string;
  delegateUpn: string;
  startUtc: string | null;
  endUtc: string | null;
  reason: string;
  isActive: boolean;
}
export interface IApprovalState {
  approvals: IApprovalRecord[];
  required: number;
  granted: number;
  satisfied: boolean;
  alreadyApprovedByMe: boolean;
}
export function isDelegationActive(delegation: IApprovalDelegation, now: number = Date.now()): boolean {
  if (!delegation.isActive || !delegation.delegateUpn) {
    return false;
  }
  const start: number | undefined = delegation.startUtc ? Date.parse(delegation.startUtc) : undefined;
  const end: number | undefined = delegation.endUtc ? Date.parse(delegation.endUtc) : undefined;
  if (start !== undefined && !isNaN(start) && now < start) {
    return false;
  }
  if (end !== undefined && !isNaN(end) && now > end) {
    return false;
  }
  return true;
}
export function evaluateApprovals(
  approvals: IApprovalRecord[],
  required: number,
  operatorUpn: string
): IApprovalState {
  const normalized: string = operatorUpn.trim().toLowerCase();
  const distinct: Set<string> = new Set(
    approvals.map((a) => (a.actorUpn ?? '').trim().toLowerCase()).filter((upn) => upn.length > 0)
  );
  const effectiveRequired: number = Math.max(1, Math.floor(required) || 1);
  return {
    approvals,
    required: effectiveRequired,
    granted: distinct.size,
    satisfied: distinct.size >= effectiveRequired,
    alreadyApprovedByMe: distinct.has(normalized)
  };
}
