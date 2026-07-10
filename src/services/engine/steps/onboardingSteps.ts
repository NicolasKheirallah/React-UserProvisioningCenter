import { validatePayload } from '../../../validators/payloadValidator';
import { generateTempPassword } from '../../passwords/passwordGenerator';
import {
  STEP_ASSIGN_APPLICATIONS,
  STEP_ASSIGN_GROUPS,
  STEP_ASSIGN_LICENSES,
  STEP_ASSIGN_MANAGER,
  STEP_ASSIGN_SHAREPOINT,
  STEP_ASSIGN_TEAMS,
  STEP_COPY_GROUPS,
  STEP_COPY_LICENSES,
  STEP_CREATE_USER,
  STEP_FINALIZE_AUDIT,
  STEP_PRESENT_CREDENTIALS,
  STEP_SCHEDULE_ACCESS_REVIEW,
  STEP_SEND_NOTIFICATIONS,
  STEP_SET_USAGE_LOCATION,
  STEP_UPLOAD_PHOTO,
  STEP_VALIDATE_CLONE_SOURCE,
  STEP_VALIDATE_INPUT
} from '../../../constants/stepIds';
import { RequestAbortedError } from '../../graph/GraphError';
import { StepFailure } from '../stepTypes';
import type { IStepContext, IWorkflowStepDefinition } from '../stepTypes';
import { auditedGraphWrite, getOrNull, runFinalizeAudit } from '../stepHelpers';
import { isOnboardingPayload } from '../../../models';
import type { IApplicationCatalogItem, IOnboardingPayload, UpnAvailability } from '../../../models';
import { escapeODataLiteral } from '../../util/odata';

/**
 * Every step is idempotent (check-before-write) so a resumed job never
 * repeats a completed side effect (spec Section 6).
 */

async function requireTargetUser(ctx: IStepContext): Promise<string> {
  if (!ctx.job.targetUserId) {
    throw new StepFailure('Target user has not been created yet', 'UPC_NoTargetUser', false);
  }
  return ctx.job.targetUserId;
}

/** Narrows the job payload; these steps only ever run for Onboard jobs. */
function onboarding(ctx: IStepContext): IOnboardingPayload {
  const payload = ctx.job.payload;
  if (!isOnboardingPayload(payload)) {
    throw new StepFailure('Job payload is not an onboarding payload', 'UPC_WrongPayloadKind', false);
  }
  return payload;
}

// ---- validate-input --------------------------------------------------------

async function runValidateInput(ctx: IStepContext): Promise<void> {
  const payload: IOnboardingPayload = onboarding(ctx);
  const errors: string[] = validatePayload(payload);
  if (errors.length > 0) {
    throw new StepFailure(`Payload validation failed: ${errors.join('; ')}`, 'UPC_InvalidPayload', false);
  }
  // Directory checks only apply while the user does not exist yet — on resume
  // after create-user we would collide with the account we just created.
  if (ctx.job.targetUserId) {
    return;
  }
  if (payload.identity.accountType === 'guest') {
    // Guests aren't subject to the employeeId/UPN acceptance gate — they're
    // not employees and Entra assigns their actual #EXT# UPN on redemption,
    // not the email address the operator entered.
    return;
  }
  const employeeIdTaken: boolean = await ctx.users.isEmployeeIdTaken(
    payload.personal.employeeId,
    ctx.signal
  );
  if (employeeIdTaken) {
    throw new StepFailure(
      `employeeId ${payload.personal.employeeId} already exists in the directory`,
      'UPC_DuplicateEmployeeId',
      false
    );
  }
  const availability: UpnAvailability = await ctx.naming.checkUpnAvailability(
    payload.identity.userPrincipalName,
    ctx.signal
  );
  if (availability !== 'available') {
    throw new StepFailure(
      `UPN ${payload.identity.userPrincipalName} is no longer available (${availability})`,
      'UPC_UpnCollision',
      false
    );
  }
}

// ---- create-user -----------------------------------------------------------

interface ICreatedUser {
  id: string;
}

async function runCreateUser(ctx: IStepContext): Promise<void> {
  const payload: IOnboardingPayload = onboarding(ctx);
  if (payload.identity.accountType === 'guest') {
    return runInviteGuest(ctx, payload);
  }

  // Idempotency 1: a previous attempt already recorded the created user.
  if (ctx.job.targetUserId) {
    const existing: ICreatedUser | null = await getOrNull(() =>
      ctx.graph.get<ICreatedUser>(`/users/${ctx.job.targetUserId}?$select=id`, { signal: ctx.signal })
    );
    if (existing) {
      return;
    }
    // Recorded id no longer resolves (deleted out-of-band) — fall through and recreate.
  }

  // Idempotency 2: the POST may have succeeded although the response was lost.
  const upn: string = payload.identity.userPrincipalName;
  const byUpn: { value: ICreatedUser[] } = await ctx.graph.get<{ value: ICreatedUser[] }>(
    `/users?$select=id&$filter=${encodeURIComponent(
      `userPrincipalName eq '${escapeODataLiteral(upn)}'`
    )}`,
    { signal: ctx.signal }
  );
  if ((byUpn.value ?? []).length > 0) {
    ctx.job.targetUserId = byUpn.value[0].id;
    await ctx.data.setJobTargetUser(ctx.job.itemId, ctx.job.targetUserId);
    return;
  }

  const password: string = generateTempPassword();
  ctx.secrets.temporaryPassword = password;

  const body: Record<string, unknown> = {
    accountEnabled: payload.accountSettings.accountEnabled,
    displayName: payload.personal.displayName,
    mailNickname: payload.identity.mailNickname,
    userPrincipalName: upn,
    passwordProfile: {
      password,
      forceChangePasswordNextSignIn: payload.accountSettings.forceChangePassword
    },
    department: payload.employment.department,
    jobTitle: payload.employment.jobTitle,
    officeLocation: payload.employment.officeLocation || undefined,
    companyName: payload.employment.companyName || undefined,
    country: payload.employment.country || undefined,
    usageLocation: payload.accountSettings.usageLocation,
    employeeId: payload.personal.employeeId,
    employeeHireDate: payload.employment.hireDate
      ? `${payload.employment.hireDate}T00:00:00Z`
      : undefined,
    employeeType: payload.employment.employeeType,
    mobilePhone: payload.personal.mobilePhone || undefined,
    otherMails: payload.personal.personalEmail ? [payload.personal.personalEmail] : undefined
  };

  const created: ICreatedUser = await auditedGraphWrite(
    ctx,
    'create-user',
    'POST',
    '/users',
    body,
    () => ctx.graph.post<ICreatedUser>('/users', body, { signal: ctx.signal })
  );
  ctx.job.targetUserId = created.id;
  await ctx.data.setJobTargetUser(ctx.job.itemId, created.id);
}

/**
 * Guest onboarding: invites an external user via /invitations instead of
 * creating a cloud account via POST /users. Entra creates the guest's
 * directory object immediately (assigning the actual #EXT# UPN) and emails
 * the invitation redemption link itself — there is no password or TAP for
 * this app to generate or hand over.
 */
async function runInviteGuest(ctx: IStepContext, payload: IOnboardingPayload): Promise<void> {
  if (ctx.job.targetUserId) {
    const existing: ICreatedUser | null = await getOrNull(() =>
      ctx.graph.get<ICreatedUser>(`/users/${ctx.job.targetUserId}?$select=id`, { signal: ctx.signal })
    );
    if (existing) {
      return;
    }
  }

  // Idempotency: the invitation may have already created the guest's
  // directory object even if the response was lost (e.g. tab closed).
  const email: string = payload.identity.userPrincipalName;
  const byMail: { value: ICreatedUser[] } = await ctx.graph.get<{ value: ICreatedUser[] }>(
    `/users?$select=id&$filter=${encodeURIComponent(`mail eq '${escapeODataLiteral(email)}'`)}`,
    { signal: ctx.signal }
  );
  if ((byMail.value ?? []).length > 0) {
    ctx.job.targetUserId = byMail.value[0].id;
    await ctx.data.setJobTargetUser(ctx.job.itemId, ctx.job.targetUserId);
    return;
  }

  const body = {
    invitedUserEmailAddress: email,
    invitedUserDisplayName: payload.personal.displayName,
    inviteRedirectUrl: 'https://myapplications.microsoft.com',
    sendInvitationMessage: true
  };
  const result: { invitedUser: ICreatedUser } = await auditedGraphWrite(
    ctx,
    'invite-guest',
    'POST',
    '/invitations',
    body,
    () => ctx.graph.post<{ invitedUser: ICreatedUser }>('/invitations', body, { signal: ctx.signal })
  );
  ctx.job.targetUserId = result.invitedUser.id;
  await ctx.data.setJobTargetUser(ctx.job.itemId, result.invitedUser.id);
}

// ---- set-usage-location ------------------------------------------------------

async function runSetUsageLocation(ctx: IStepContext): Promise<void> {
  if (onboarding(ctx).identity.accountType === 'guest') {
    return; // usage location gates license assignment, which doesn't apply to guests here
  }
  const userId: string = await requireTargetUser(ctx);
  const desired: string = onboarding(ctx).accountSettings.usageLocation;
  const current: { usageLocation: string | null } | null = await getOrNull(() =>
    ctx.graph.get<{ usageLocation: string | null }>(`/users/${userId}?$select=usageLocation`, {
      signal: ctx.signal
    })
  );
  if (current?.usageLocation === desired) {
    return;
  }
  const body = { usageLocation: desired };
  await auditedGraphWrite(ctx, 'set-usage-location', 'PATCH', `/users/${userId}`, body, () =>
    ctx.graph.patch<void>(`/users/${userId}`, body, { signal: ctx.signal })
  );
}

// ---- assign-manager ----------------------------------------------------------

async function runAssignManager(ctx: IStepContext): Promise<void> {
  const managerId: string | undefined = onboarding(ctx).employment.managerId;
  if (!managerId) {
    return; // no manager selected — legitimate no-op
  }
  const userId: string = await requireTargetUser(ctx);
  const currentManager: { id: string } | null = await getOrNull(() =>
    ctx.graph.get<{ id: string }>(`/users/${userId}/manager?$select=id`, { signal: ctx.signal })
  );
  if (currentManager?.id === managerId) {
    return;
  }
  const body = { '@odata.id': `https://graph.microsoft.com/v1.0/users/${managerId}` };
  await auditedGraphWrite(
    ctx,
    'assign-manager',
    'PUT',
    `/users/${userId}/manager/$ref`,
    body,
    () => ctx.graph.put<void>(`/users/${userId}/manager/$ref`, body, { signal: ctx.signal })
  );
}

// ---- assign-licenses ---------------------------------------------------------

async function runAssignLicenses(ctx: IStepContext): Promise<void> {
  const payload = onboarding(ctx);
  if (payload.identity.accountType === 'guest') {
    return; // guests in this flow are not licensed
  }
  const selections = payload.licenses;
  if (selections.length === 0) {
    return;
  }
  const userId: string = await requireTargetUser(ctx);
  const details: { value: { skuId: string }[] } = await ctx.graph.get<{
    value: { skuId: string }[];
  }>(`/users/${userId}/licenseDetails?$select=skuId`, { signal: ctx.signal });
  const already: Set<string> = new Set((details.value ?? []).map((d) => d.skuId));
  const toAdd: string[] = selections.map((s) => s.skuId).filter((skuId) => !already.has(skuId));
  if (toAdd.length === 0) {
    return;
  }
  const body = {
    addLicenses: toAdd.map((skuId) => ({ skuId, disabledPlans: [] })),
    removeLicenses: []
  };
  await auditedGraphWrite(
    ctx,
    'assign-licenses',
    'POST',
    `/users/${userId}/assignLicense`,
    body,
    () => ctx.graph.post<void>(`/users/${userId}/assignLicense`, body, { signal: ctx.signal })
  );
}

// ---- shared group-membership helpers (assign-groups + assign-applications) ----

/** One batched checkMemberGroups call instead of one GET per candidate group. */
async function alreadyMemberOf(
  ctx: IStepContext,
  userId: string,
  groupIds: string[]
): Promise<Set<string>> {
  if (groupIds.length === 0) {
    return new Set();
  }
  const result: { value: string[] } | null = await getOrNull(() =>
    ctx.graph.post<{ value: string[] }>(
      `/users/${userId}/checkMemberGroups`,
      { groupIds },
      { signal: ctx.signal }
    )
  );
  return new Set(result?.value ?? []);
}

async function addGroupMember(ctx: IStepContext, groupId: string, userId: string, action: string): Promise<void> {
  const body = { '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}` };
  await auditedGraphWrite(ctx, action, 'POST', `/groups/${groupId}/members/$ref`, body, () =>
    ctx.graph.post<void>(`/groups/${groupId}/members/$ref`, body, { signal: ctx.signal })
  );
}

// ---- assign-groups -----------------------------------------------------------

async function runAssignGroups(ctx: IStepContext): Promise<void> {
  const payload = onboarding(ctx);
  const groupIds: string[] = [...payload.access.securityGroups, ...payload.access.m365Groups];
  if (groupIds.length === 0) {
    return;
  }
  const userId: string = await requireTargetUser(ctx);
  const already: Set<string> = await alreadyMemberOf(ctx, userId, groupIds);

  const failed: string[] = [];
  for (const groupId of groupIds) {
    if (already.has(groupId)) {
      continue;
    }
    try {
      await addGroupMember(ctx, groupId, userId, 'assign-group');
    } catch (err) {
      if (err instanceof RequestAbortedError) {
        throw err;
      }
      // Already audited as a failure by auditedGraphWrite; hand it to the desk.
      failed.push(groupId);
    }
  }
  if (failed.length > 0) {
    await ctx.data.createTask(
      ctx.job.jobId,
      'GroupAssignment',
      `Add ${payload.identity.userPrincipalName} to remaining groups`,
      `These group memberships could not be added through Graph: ${failed.join(', ')}`
    );
  }
}

// ---- assign-teams -----------------------------------------------------------

async function runAssignTeams(ctx: IStepContext): Promise<void> {
  const payload = onboarding(ctx);
  const assignments = payload.access.teams;
  if (assignments.length === 0) {
    return;
  }
  const userId: string = await requireTargetUser(ctx);

  const failed: string[] = [];
  for (const assignment of assignments) {
    const existing: { value: { userId?: string }[] } | null = await getOrNull(() =>
      ctx.graph.get<{ value: { userId?: string }[] }>(`/teams/${assignment.teamId}/members?$top=999`, {
        signal: ctx.signal
      })
    );
    if ((existing?.value ?? []).some((m) => m.userId === userId)) {
      continue;
    }
    const body = {
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: assignment.role === 'owner' ? ['owner'] : [],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`
    };
    try {
      await auditedGraphWrite(ctx, 'assign-team', 'POST', `/teams/${assignment.teamId}/members`, body, () =>
        ctx.graph.post<void>(`/teams/${assignment.teamId}/members`, body, { signal: ctx.signal })
      );
    } catch (err) {
      if (err instanceof RequestAbortedError) {
        throw err;
      }
      failed.push(assignment.teamId);
    }
  }
  if (failed.length > 0) {
    await ctx.data.createTask(
      ctx.job.jobId,
      'TeamAssignment',
      `Add ${payload.identity.userPrincipalName} to remaining Teams`,
      `These Teams memberships could not be added through Graph: ${failed.join(', ')}`
    );
  }
}

// ---- assign-sharepoint --------------------------------------------------------

async function runAssignSharePoint(ctx: IStepContext): Promise<void> {
  const payload = onboarding(ctx);
  const assignments = payload.access.sharePointSites;
  if (assignments.length === 0) {
    return;
  }
  await requireTargetUser(ctx);

  const failed: string[] = [];
  for (const assignment of assignments) {
    const started: number = Date.now();
    // Not a Graph call — SiteAccessService uses the operator's own SharePoint
    // permissions on the target site via PnPJS (see its class doc comment).
    // Logged directly rather than through auditedGraphWrite, which assumes a
    // Graph endpoint/response code.
    try {
      await ctx.siteAccess.grantAccess(assignment.siteUrl, payload.identity.userPrincipalName, assignment.role);
      await ctx.audit.log({
        jobId: ctx.job.jobId,
        action: 'assign-sharepoint-access',
        targetUser: payload.identity.userPrincipalName,
        graphEndpoint: '',
        requestSummary: JSON.stringify({ siteUrl: assignment.siteUrl, role: assignment.role }),
        responseCode: 200,
        durationMs: Date.now() - started,
        result: 'Success',
        correlationId: ctx.job.correlationId
      });
    } catch {
      await ctx.audit.log({
        jobId: ctx.job.jobId,
        action: 'assign-sharepoint-access',
        targetUser: payload.identity.userPrincipalName,
        graphEndpoint: '',
        requestSummary: JSON.stringify({ siteUrl: assignment.siteUrl, role: assignment.role }),
        responseCode: 0,
        durationMs: Date.now() - started,
        result: 'Failure',
        correlationId: ctx.job.correlationId
      });
      failed.push(assignment.siteUrl);
    }
  }
  if (failed.length > 0) {
    await ctx.data.createTask(
      ctx.job.jobId,
      'SharePointAccess',
      `Grant ${payload.identity.userPrincipalName} SharePoint access`,
      `These site grants could not be completed automatically: ${failed.join(', ')}`
    );
  }
}

// ---- assign-applications ------------------------------------------------------

async function runAssignApplications(ctx: IStepContext): Promise<void> {
  const payload = onboarding(ctx);
  const appItemIds = payload.access.applications;
  if (appItemIds.length === 0) {
    return;
  }
  const userId: string = await requireTargetUser(ctx);

  const catalog: IApplicationCatalogItem[] = await ctx.data.getApplicationCatalog();
  const byId: Map<string, IApplicationCatalogItem> = new Map(
    catalog.map((a) => [String(a.itemId), a])
  );

  const groupBasedTargets: { app: IApplicationCatalogItem }[] = [];
  const manualEntries: string[] = [];
  for (const appItemId of appItemIds) {
    const app = byId.get(appItemId);
    if (!app) {
      continue; // catalog entry removed/deactivated since the wizard/template selected it
    }
    if (app.provisioningType === 'GroupBased' && app.targetGroupId) {
      groupBasedTargets.push({ app });
    } else {
      manualEntries.push(app.title);
    }
  }

  const groupIds: string[] = groupBasedTargets.map((t) => t.app.targetGroupId as string);
  const already: Set<string> = await alreadyMemberOf(ctx, userId, groupIds);

  const failedGroupBased: string[] = [];
  for (const { app } of groupBasedTargets) {
    const groupId: string = app.targetGroupId as string;
    if (already.has(groupId)) {
      continue;
    }
    try {
      await addGroupMember(ctx, groupId, userId, 'assign-application');
    } catch (err) {
      if (err instanceof RequestAbortedError) {
        throw err;
      }
      failedGroupBased.push(app.title);
    }
  }

  const needsAttention: string[] = [...manualEntries, ...failedGroupBased];
  if (needsAttention.length > 0) {
    await ctx.data.createTask(
      ctx.job.jobId,
      'ApplicationAssignment',
      `Grant ${payload.identity.userPrincipalName} access to applications`,
      `These applications need manual provisioning or could not be group-assigned: ${needsAttention.join(', ')}`
    );
  }
}

// ---- present-credentials ---------------------------------------------------------
// There is no wait-for-mailbox step: creating the user does not provision a
// mailbox — that happens asynchronously once assign-licenses succeeds, on
// Exchange's own timeline, with no API to poll for readiness against the
// permission scopes this app holds. Nothing downstream (this step included)
// depends on the mailbox existing, so there is nothing to gate on here.

async function runPresentCredentials(ctx: IStepContext): Promise<void> {
  const payload: IOnboardingPayload = onboarding(ctx);
  if (payload.identity.accountType === 'guest') {
    // Entra emails the invitation redemption link directly — there is no
    // password or TAP for this app to generate or hand over for a guest.
    return;
  }
  const userId: string = await requireTargetUser(ctx);
  const upn: string = payload.identity.userPrincipalName;

  if (payload.accountSettings.credentialMode === 'tap') {
    // A user can hold only one TAP: drop any stale one before creating ours
    // (its value is unreadable after creation, so it cannot be re-shown).
    const existing: { value: { id: string }[] } | null = await getOrNull(() =>
      ctx.graph.get<{ value: { id: string }[] }>(
        `/users/${userId}/authentication/temporaryAccessPassMethods`,
        { signal: ctx.signal }
      )
    );
    for (const method of existing?.value ?? []) {
      await auditedGraphWrite(
        ctx,
        'delete-stale-tap',
        'DELETE',
        `/users/${userId}/authentication/temporaryAccessPassMethods/${method.id}`,
        undefined,
        () =>
          ctx.graph.delete<void>(
            `/users/${userId}/authentication/temporaryAccessPassMethods/${method.id}`,
            { signal: ctx.signal }
          )
      );
    }
    const tap: { temporaryAccessPass: string } = await auditedGraphWrite(
      ctx,
      'create-tap',
      'POST',
      `/users/${userId}/authentication/temporaryAccessPassMethods`,
      {},
      () =>
        ctx.graph.post<{ temporaryAccessPass: string }>(
          `/users/${userId}/authentication/temporaryAccessPassMethods`,
          {},
          { signal: ctx.signal }
        )
    );
    ctx.secrets.temporaryAccessPass = tap.temporaryAccessPass;
    await ctx.presentCredentials({
      kind: 'tap',
      value: tap.temporaryAccessPass,
      userPrincipalName: upn
    });
    return;
  }

  // Password mode. If the in-memory password from create-user is gone (the
  // job resumed in a new session), reset it — the old value is unrecoverable
  // by design (never persisted anywhere).
  if (!ctx.secrets.temporaryPassword) {
    const password: string = generateTempPassword();
    const body = {
      passwordProfile: {
        password,
        forceChangePasswordNextSignIn: payload.accountSettings.forceChangePassword
      }
    };
    await auditedGraphWrite(ctx, 'reset-temp-password', 'PATCH', `/users/${userId}`, body, () =>
      ctx.graph.patch<void>(`/users/${userId}`, body, { signal: ctx.signal })
    );
    ctx.secrets.temporaryPassword = password;
  }
  await ctx.presentCredentials({
    kind: 'password',
    value: ctx.secrets.temporaryPassword,
    userPrincipalName: upn
  });
}

// ---- upload-photo --------------------------------------------------------------

async function runUploadPhoto(ctx: IStepContext): Promise<void> {
  const payload = onboarding(ctx);
  const dataUrl: string | undefined = payload.personal.photoDataUrl;
  if (!dataUrl) {
    return;
  }
  const userId: string = await requireTargetUser(ctx);

  const match: RegExpExecArray | null = /^data:([^;]+);base64,(.+)$/.exec(dataUrl);
  if (!match) {
    return; // malformed data URL — nothing sensible to upload, not worth failing the job over
  }
  const [, contentType, base64] = match;
  const binary: string = atob(base64);
  const bytes: Uint8Array = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }

  await auditedGraphWrite(
    ctx,
    'upload-photo',
    'PUT',
    `/users/${userId}/photo/$value`,
    { contentType, bytes: bytes.length }, // requestSummary only — never the binary itself
    () =>
      ctx.graph.put<void>(`/users/${userId}/photo/$value`, bytes.buffer, {
        signal: ctx.signal,
        headers: { 'Content-Type': contentType }
      })
  );
}

// ---- send-notifications ---------------------------------------------------------

async function runSendNotifications(ctx: IStepContext): Promise<void> {
  const payload = onboarding(ctx);
  const managerUpn: string | undefined = payload.employment.managerUpn;
  if (!managerUpn) {
    return; // no manager selected — nobody to notify
  }
  const body = {
    message: {
      subject: `Onboarding complete: ${payload.personal.displayName}`,
      body: {
        contentType: 'Text',
        content:
          `${payload.personal.displayName}'s account (${payload.identity.userPrincipalName}) ` +
          'has been provisioned and is ready for first sign-in.'
      },
      toRecipients: [{ emailAddress: { address: managerUpn } }]
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

// ---- schedule-access-review ------------------------------------------------------

async function runScheduleAccessReview(ctx: IStepContext): Promise<void> {
  const payload = onboarding(ctx);
  const days: number | null = payload.expirationReviewDays;
  if (!days || days <= 0) {
    return;
  }
  // No server exists to act on this unattended — the review is a task for a
  // human, not an automatic future action (spec: expiration policy).
  const dueDateLabel: string = new Date(Date.now() + days * 24 * 60 * 60 * 1000)
    .toISOString()
    .slice(0, 10);
  await ctx.data.createTask(
    ctx.job.jobId,
    'AccessReview',
    `Review access for ${payload.identity.userPrincipalName} by ${dueDateLabel}`,
    `This account's access was granted under a ${days}-day review policy. Review whether ` +
      `${payload.identity.userPrincipalName}'s licenses, groups and access are still needed, ` +
      `and offboard or renew as appropriate. Due: ${dueDateLabel}.`
  );
}

// ---- clone-specific steps (validate-clone-source, copy-licenses, copy-groups) ----
// Clone reuses this entire pipeline (see CLONE_STEP_ORDER in stepIds.ts) with
// these three steps inserted after assign-licenses — the cloned user still
// goes through create-user, its own template selections, credentials, etc.,
// on top of what's copied from the source.

function cloneSource(ctx: IStepContext): string | undefined {
  return onboarding(ctx).cloneSourceUserId;
}

async function runValidateCloneSource(ctx: IStepContext): Promise<void> {
  const sourceId: string | undefined = cloneSource(ctx);
  if (!sourceId) {
    throw new StepFailure('Clone job has no source user set', 'UPC_InvalidPayload', false);
  }
  const source: { id: string } | null = await getOrNull(() =>
    ctx.graph.get<{ id: string }>(`/users/${sourceId}?$select=id`, { signal: ctx.signal })
  );
  if (!source) {
    throw new StepFailure('The clone source user was not found in the directory', 'UPC_TargetNotFound', false);
  }
}

async function runCopyLicenses(ctx: IStepContext): Promise<void> {
  const sourceId: string | undefined = cloneSource(ctx);
  if (!sourceId) {
    return;
  }
  const userId: string = await requireTargetUser(ctx);
  const source: { value: { skuId: string }[] } = await ctx.graph.get<{ value: { skuId: string }[] }>(
    `/users/${sourceId}/licenseDetails?$select=skuId`,
    { signal: ctx.signal }
  );
  const sourceSkuIds: string[] = (source.value ?? []).map((d) => d.skuId);
  if (sourceSkuIds.length === 0) {
    return;
  }
  const target: { value: { skuId: string }[] } = await ctx.graph.get<{ value: { skuId: string }[] }>(
    `/users/${userId}/licenseDetails?$select=skuId`,
    { signal: ctx.signal }
  );
  const already: Set<string> = new Set((target.value ?? []).map((d) => d.skuId));
  const toAdd: string[] = sourceSkuIds.filter((id) => !already.has(id));
  if (toAdd.length === 0) {
    return;
  }
  const body = { addLicenses: toAdd.map((skuId) => ({ skuId, disabledPlans: [] })), removeLicenses: [] };
  await auditedGraphWrite(ctx, 'copy-licenses', 'POST', `/users/${userId}/assignLicense`, body, () =>
    ctx.graph.post<void>(`/users/${userId}/assignLicense`, body, { signal: ctx.signal })
  );
}

interface ISourceGroup {
  id: string;
  groupTypes?: string[];
  mailEnabled?: boolean;
  onPremisesSyncEnabled?: boolean | null;
}
interface ISourceGroupPage {
  value: ISourceGroup[];
  '@odata.nextLink'?: string;
}

/** Same exclusions as offboarding's group removal: dynamic/synced/DL-type
 *  memberships can't be edited through Graph in either direction. */
async function fetchAssignableGroups(ctx: IStepContext, sourceUserId: string): Promise<string[]> {
  const groups: ISourceGroup[] = [];
  let url: string | undefined =
    `/users/${sourceUserId}/memberOf/microsoft.graph.group` +
    '?$select=id,groupTypes,mailEnabled,onPremisesSyncEnabled&$top=999';
  while (url) {
    const page: ISourceGroupPage = await ctx.graph.get<ISourceGroupPage>(url, { signal: ctx.signal });
    groups.push(...(page.value ?? []));
    url = page['@odata.nextLink'];
  }
  return groups
    .filter((g) => {
      const types: string[] = g.groupTypes ?? [];
      const isDynamic: boolean = types.indexOf('DynamicMembership') !== -1;
      const isUnified: boolean = types.indexOf('Unified') !== -1;
      return !isDynamic && !g.onPremisesSyncEnabled && (!g.mailEnabled || isUnified);
    })
    .map((g) => g.id);
}

async function runCopyGroups(ctx: IStepContext): Promise<void> {
  const sourceId: string | undefined = cloneSource(ctx);
  if (!sourceId) {
    return;
  }
  const userId: string = await requireTargetUser(ctx);
  const groupIds: string[] = await fetchAssignableGroups(ctx, sourceId);
  if (groupIds.length === 0) {
    return;
  }
  const already: Set<string> = await alreadyMemberOf(ctx, userId, groupIds);

  const failed: string[] = [];
  for (const groupId of groupIds) {
    if (already.has(groupId)) {
      continue;
    }
    try {
      await addGroupMember(ctx, groupId, userId, 'copy-group');
    } catch (err) {
      if (err instanceof RequestAbortedError) {
        throw err;
      }
      failed.push(groupId);
    }
  }
  if (failed.length > 0) {
    await ctx.data.createTask(
      ctx.job.jobId,
      'GroupAssignment',
      `Add ${onboarding(ctx).identity.userPrincipalName} to remaining cloned groups`,
      `These group memberships could not be copied through Graph: ${failed.join(', ')}`
    );
  }
}

// ---- registry -------------------------------------------------------------------
// runFinalizeAudit is shared with offboardingSteps.ts/transferSteps.ts — see stepHelpers.ts.

export const ONBOARDING_STEPS: IWorkflowStepDefinition[] = [
  { id: STEP_VALIDATE_INPUT, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runValidateInput },
  { id: STEP_CREATE_USER, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runCreateUser },
  { id: STEP_SET_USAGE_LOCATION, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runSetUsageLocation },
  { id: STEP_ASSIGN_MANAGER, skippable: true, maxAttempts: 3, continueOnFailure: false, run: runAssignManager },
  { id: STEP_ASSIGN_LICENSES, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runAssignLicenses },
  { id: STEP_ASSIGN_GROUPS, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runAssignGroups },
  { id: STEP_ASSIGN_TEAMS, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runAssignTeams },
  { id: STEP_ASSIGN_SHAREPOINT, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runAssignSharePoint },
  { id: STEP_ASSIGN_APPLICATIONS, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runAssignApplications },
  { id: STEP_UPLOAD_PHOTO, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runUploadPhoto },
  { id: STEP_PRESENT_CREDENTIALS, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runPresentCredentials },
  { id: STEP_SEND_NOTIFICATIONS, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runSendNotifications },
  { id: STEP_SCHEDULE_ACCESS_REVIEW, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runScheduleAccessReview },
  { id: STEP_FINALIZE_AUDIT, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runFinalizeAudit }
];

export const CLONE_STEPS: IWorkflowStepDefinition[] = [
  { id: STEP_VALIDATE_INPUT, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runValidateInput },
  { id: STEP_CREATE_USER, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runCreateUser },
  { id: STEP_SET_USAGE_LOCATION, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runSetUsageLocation },
  { id: STEP_ASSIGN_MANAGER, skippable: true, maxAttempts: 3, continueOnFailure: false, run: runAssignManager },
  { id: STEP_ASSIGN_LICENSES, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runAssignLicenses },
  { id: STEP_VALIDATE_CLONE_SOURCE, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runValidateCloneSource },
  { id: STEP_COPY_LICENSES, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runCopyLicenses },
  { id: STEP_COPY_GROUPS, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runCopyGroups },
  { id: STEP_ASSIGN_GROUPS, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runAssignGroups },
  { id: STEP_ASSIGN_TEAMS, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runAssignTeams },
  { id: STEP_ASSIGN_SHAREPOINT, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runAssignSharePoint },
  { id: STEP_ASSIGN_APPLICATIONS, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runAssignApplications },
  { id: STEP_UPLOAD_PHOTO, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runUploadPhoto },
  { id: STEP_PRESENT_CREDENTIALS, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runPresentCredentials },
  { id: STEP_SEND_NOTIFICATIONS, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runSendNotifications },
  { id: STEP_SCHEDULE_ACCESS_REVIEW, skippable: true, maxAttempts: 3, continueOnFailure: true, run: runScheduleAccessReview },
  { id: STEP_FINALIZE_AUDIT, skippable: false, maxAttempts: 3, continueOnFailure: false, run: runFinalizeAudit }
];
