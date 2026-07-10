import {
  STEP_BLOCK_SIGN_IN,
  STEP_CREATE_HANDOVER_TASKS,
  STEP_FINALIZE_AUDIT,
  STEP_REMOVE_GROUPS,
  STEP_REMOVE_LICENSES,
  STEP_SEND_NOTIFICATIONS,
  STEP_VALIDATE_TARGET
} from '../../../constants/stepIds';
import { RequestAbortedError } from '../../graph/GraphError';
import { StepFailure } from '../stepTypes';
import type { IStepContext, IWorkflowStepDefinition } from '../stepTypes';
import { auditedGraphWrite, getOrNull, runFinalizeAudit } from '../stepHelpers';
import { isOffboardingPayload } from '../../../models';
import type { IOffboardingPayload } from '../../../models';

/**
 * Offboarding pipeline (spec Sections 6 and 10). Every step is idempotent;
 * anything Graph cannot perform (shared-mailbox conversion, mail forwarding,
 * OneDrive hand-over, Exchange-managed group removal) lands in UPC_Tasks for
 * the service desk instead of failing the job.
 */

function offboarding(ctx: IStepContext): IOffboardingPayload {
  const payload = ctx.job.payload;
  if (!isOffboardingPayload(payload)) {
    throw new StepFailure('Job payload is not an offboarding payload', 'UPC_WrongPayloadKind', false);
  }
  return payload;
}

// ---- validate-target ---------------------------------------------------------

async function runValidateTarget(ctx: IStepContext): Promise<void> {
  const payload: IOffboardingPayload = offboarding(ctx);
  const target = payload.target;
  if (!target?.userId || !target.userPrincipalName) {
    throw new StepFailure('Offboarding payload has no target user', 'UPC_InvalidPayload', false);
  }
  if (payload.options.mailboxAction === 'forward' && !payload.options.forwardingAddress) {
    throw new StepFailure(
      'Mail forwarding selected but no forwarding address provided',
      'UPC_InvalidPayload',
      false
    );
  }
  const user: { id: string } | null = await getOrNull(() =>
    ctx.graph.get<{ id: string }>(`/users/${target.userId}?$select=id`, { signal: ctx.signal })
  );
  if (!user) {
    throw new StepFailure(
      `User ${target.userPrincipalName} was not found in the directory`,
      'UPC_TargetNotFound',
      false
    );
  }
  if (!ctx.job.targetUserId) {
    ctx.job.targetUserId = target.userId;
    await ctx.data.setJobTargetUser(ctx.job.itemId, target.userId);
  }
}

// ---- block-sign-in -----------------------------------------------------------

async function runBlockSignIn(ctx: IStepContext): Promise<void> {
  const payload: IOffboardingPayload = offboarding(ctx);
  const userId: string = payload.target.userId;

  const current: { accountEnabled: boolean; onPremisesSyncEnabled: boolean | null } | null =
    await getOrNull(() =>
      ctx.graph.get<{ accountEnabled: boolean; onPremisesSyncEnabled: boolean | null }>(
        `/users/${userId}?$select=accountEnabled,onPremisesSyncEnabled`,
        { signal: ctx.signal }
      )
    );

  if (current?.onPremisesSyncEnabled) {
    // accountEnabled is a directory-synced attribute for hybrid users — its
    // source of truth is on-prem AD. A cloud-side Graph write here would
    // normally be overwritten by the next AAD Connect sync cycle (on-prem
    // wins), so the job would silently "un-disable" the account later while
    // reporting success today. Hand off to whoever manages on-prem AD
    // instead of writing a change that won't survive the next sync.
    await ctx.data.createTask(
      ctx.job.jobId,
      'OnPremAdAccount',
      `Disable on-premises AD account: ${payload.target.userPrincipalName}`,
      `${payload.target.userPrincipalName} is synced from on-premises Active Directory ` +
        '(onPremisesSyncEnabled). Disabling sign-in through Graph would be reverted by the ' +
        'next directory sync — disable or move the account in on-prem AD instead so the ' +
        'change survives sync.'
    );
  } else if (current?.accountEnabled !== false) {
    const body = { accountEnabled: false };
    await auditedGraphWrite(ctx, 'block-sign-in', 'PATCH', `/users/${userId}`, body, () =>
      ctx.graph.patch<void>(`/users/${userId}`, body, { signal: ctx.signal })
    );
  }
  // Session revocation is an action, not a synced-attribute write — it takes
  // effect immediately in Entra regardless of hybrid sync status, and is
  // safe to repeat on resume since it isn't observable beforehand.
  await auditedGraphWrite(
    ctx,
    'revoke-sessions',
    'POST',
    `/users/${userId}/revokeSignInSessions`,
    {},
    () => ctx.graph.post<void>(`/users/${userId}/revokeSignInSessions`, {}, { signal: ctx.signal })
  );
}

// ---- remove-licenses ---------------------------------------------------------

async function runRemoveLicenses(ctx: IStepContext): Promise<void> {
  const payload: IOffboardingPayload = offboarding(ctx);
  if (!payload.options.removeLicenses) {
    return;
  }
  const userId: string = payload.target.userId;
  const details: { value: { skuId: string }[] } = await ctx.graph.get<{
    value: { skuId: string }[];
  }>(`/users/${userId}/licenseDetails?$select=skuId`, { signal: ctx.signal });
  const skuIds: string[] = (details.value ?? []).map((d) => d.skuId);
  if (skuIds.length === 0) {
    return;
  }
  const body = { addLicenses: [], removeLicenses: skuIds };
  await auditedGraphWrite(
    ctx,
    'remove-licenses',
    'POST',
    `/users/${userId}/assignLicense`,
    body,
    () => ctx.graph.post<void>(`/users/${userId}/assignLicense`, body, { signal: ctx.signal })
  );
}

// ---- remove-groups -----------------------------------------------------------

interface IMemberGroup {
  id: string;
  displayName: string;
  groupTypes?: string[];
  mailEnabled?: boolean;
  onPremisesSyncEnabled?: boolean | null;
}

interface IMemberGroupPage {
  value: IMemberGroup[];
  '@odata.nextLink'?: string;
}

/** $top=999 is the page ceiling for this endpoint; follow @odata.nextLink so a
 *  user in more than 999 groups doesn't silently have the remainder skipped. */
async function fetchAllMemberGroups(ctx: IStepContext, userId: string): Promise<IMemberGroup[]> {
  const groups: IMemberGroup[] = [];
  let url: string | undefined =
    `/users/${userId}/memberOf/microsoft.graph.group` +
    '?$select=id,displayName,groupTypes,mailEnabled,onPremisesSyncEnabled&$top=999';
  while (url) {
    const page: IMemberGroupPage = await ctx.graph.get<IMemberGroupPage>(url, { signal: ctx.signal });
    groups.push(...(page.value ?? []));
    url = page['@odata.nextLink'];
  }
  return groups;
}

async function runRemoveGroups(ctx: IStepContext): Promise<void> {
  const payload: IOffboardingPayload = offboarding(ctx);
  if (!payload.options.removeFromGroups) {
    return;
  }
  const userId: string = payload.target.userId;
  const allGroups: IMemberGroup[] = await fetchAllMemberGroups(ctx, userId);

  const notRemovable: string[] = [];
  for (const group of allGroups) {
    const groupTypes: string[] = group.groupTypes ?? [];
    const isDynamic: boolean = groupTypes.indexOf('DynamicMembership') !== -1;
    const isUnified: boolean = groupTypes.indexOf('Unified') !== -1;
    // Distribution lists / mail-enabled security groups are Exchange-managed;
    // dynamic and on-prem-synced memberships cannot be edited through Graph.
    if (isDynamic || group.onPremisesSyncEnabled || (group.mailEnabled && !isUnified)) {
      notRemovable.push(group.displayName);
      continue;
    }
    try {
      await auditedGraphWrite(
        ctx,
        'remove-group-member',
        'DELETE',
        `/groups/${group.id}/members/${userId}/$ref`,
        undefined,
        () =>
          ctx.graph.delete<void>(`/groups/${group.id}/members/${userId}/$ref`, {
            signal: ctx.signal
          })
      );
    } catch (err) {
      if (err instanceof RequestAbortedError) {
        throw err;
      }
      // Already audited as a failure by auditedGraphWrite; hand it to the desk.
      notRemovable.push(group.displayName);
    }
  }

  if (notRemovable.length > 0) {
    await ctx.data.createTask(
      ctx.job.jobId,
      'DistributionList',
      `Remove ${payload.target.userPrincipalName} from remaining groups`,
      `These memberships could not be removed through Graph: ${notRemovable.join(', ')}`
    );
  }
}

// ---- create-handover-tasks -----------------------------------------------------

async function runCreateHandoverTasks(ctx: IStepContext): Promise<void> {
  const payload: IOffboardingPayload = offboarding(ctx);
  const upn: string = payload.target.userPrincipalName;
  const options = payload.options;
  // A retry after a mid-step failure can duplicate a task; the desk treats
  // duplicates as no-ops, which beats silently dropping a hand-over.
  if (options.mailboxAction === 'convertShared') {
    await ctx.data.createTask(
      ctx.job.jobId,
      'MailboxConvertShared',
      `Convert mailbox to shared: ${upn}`,
      `Convert the mailbox of ${upn} to a shared mailbox (Exchange admin center or PowerShell).`
    );
  }
  if (options.mailboxAction === 'forward' && options.forwardingAddress) {
    await ctx.data.createTask(
      ctx.job.jobId,
      'MailForwarding',
      `Set mail forwarding: ${upn}`,
      `Forward mail from ${upn} to ${options.forwardingAddress}.`
    );
  }
  if (options.oneDriveAccessUpn) {
    await ctx.data.createTask(
      ctx.job.jobId,
      'OneDriveHandling',
      `Grant OneDrive access: ${upn}`,
      `Grant ${options.oneDriveAccessUpn} access to the OneDrive content of ${upn}.`
    );
  }
}

// ---- send-notifications --------------------------------------------------------

async function runSendNotifications(ctx: IStepContext): Promise<void> {
  const payload: IOffboardingPayload = offboarding(ctx);
  // oneDriveAccessUpn doubles as "who's handling the hand-over" — usually the
  // manager. No manager/hand-over contact on file means nobody to notify.
  const notifyUpn: string | undefined = payload.options.oneDriveAccessUpn;
  if (!notifyUpn) {
    return;
  }
  const body = {
    message: {
      subject: `Offboarding complete: ${payload.target.displayName}`,
      body: {
        contentType: 'Text',
        content:
          `${payload.target.displayName}'s account (${payload.target.userPrincipalName}) ` +
          'has been offboarded. Sign-in is blocked and active sessions have been revoked.'
      },
      toRecipients: [{ emailAddress: { address: notifyUpn } }]
    },
    saveToSentItems: true
  };
  // Safe to repeat on resume — like revoke-sessions, sending mail isn't
  // observable beforehand, so a rare duplicate notification on a resumed
  // job is an acceptable trade-off against skipping it outright.
  await auditedGraphWrite(ctx, 'send-notification', 'POST', '/me/sendMail', body, () =>
    ctx.graph.post<void>('/me/sendMail', body, { signal: ctx.signal })
  );
}

// ---- registry --------------------------------------------------------------------
// runFinalizeAudit is shared with onboardingSteps.ts — see stepHelpers.ts.

export const OFFBOARDING_STEPS: IWorkflowStepDefinition[] = [
  { id: STEP_VALIDATE_TARGET, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runValidateTarget },
  { id: STEP_BLOCK_SIGN_IN, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runBlockSignIn },
  { id: STEP_REMOVE_LICENSES, skippable: true, maxAttempts: 3, continueOnFailure: false, run: runRemoveLicenses },
  { id: STEP_REMOVE_GROUPS, skippable: true, maxAttempts: 3, continueOnFailure: false, run: runRemoveGroups },
  { id: STEP_CREATE_HANDOVER_TASKS, skippable: true, maxAttempts: 3, continueOnFailure: false, run: runCreateHandoverTasks },
  { id: STEP_SEND_NOTIFICATIONS, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runSendNotifications },
  { id: STEP_FINALIZE_AUDIT, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runFinalizeAudit }
];
