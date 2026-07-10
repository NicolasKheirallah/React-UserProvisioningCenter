# User Provisioning Center — Permission Model & Deployment

SPFx web part for employee onboarding, lifecycle management and offboarding.
**Client-only architecture: no Azure Functions, no custom backend.** Every
operation runs in the operator's browser through **delegated** Microsoft Graph
permissions (`msGraphClientFactory.getClient("3")`).

## How security actually works here

With delegated permissions, effective access = granted app scope ∩ the
signed-in operator's own Entra privileges. A user without the relevant Entra
directory role cannot create users or assign licenses through this web part,
no matter which scopes were approved.

**Enforcement lives in Entra directory roles, not in this app.** The app's
role model (`UPC_Roles`) only controls what the UI shows and offers — it is a
usability layer, never a security boundary. On load the app runs a permission
preflight and shows a MessageBar listing capabilities the current operator
lacks.

## Operators must hold real Entra roles

| Capability | Minimum Entra role for the operator |
| --- | --- |
| Create/update users, assign manager | User Administrator |
| Assign/remove licenses | License Administrator |
| Group membership writes | Groups Administrator (or group owner) |
| Teams membership writes | Teams Administrator (or team owner) |
| Temporary Access Pass creation | Authentication Administrator |
| Guest invitations | Guest Inviter (or User Administrator) |
| Revoke sign-in sessions (offboarding) | User Administrator |

## Requested Graph scopes and why

Trimmed to exactly what the app calls today — a scope is only requested if a
Graph call in `src/services/**` actually exercises it. `PreflightService`
evaluates whether the operator holds the Teams Administrator and Guest
Inviter/User Administrator directory roles (see the table above) by reading
`/me/memberOf` (covered by `Directory.Read.All`).

| Scope | Justification |
| --- | --- |
| `User.ReadWrite.All` | Create users, set attributes/usage location, assign manager, reset first-sign-in password, revoke sessions at offboarding, Transfer field/manager updates. |
| `GroupMember.ReadWrite.All` | Group membership writes only (add/remove members) for onboarding access grants, application group-based provisioning, clone, and offboarding cleanup — not full group object create/update/delete. |
| `Organization.Read.All` | Read subscribed SKUs (license availability) and verified domains. |
| `Directory.Read.All` | UPN/mail/proxyAddresses collision checks incl. soft-deleted users; operator directory-role preflight; role resolution via group membership; the Access step's security/M365 group search-and-select picker. |
| `UserAuthenticationMethod.ReadWrite.All` | Create/regenerate Temporary Access Passes for first sign-in and credential regeneration. |
| `Mail.Send` | Notify a new hire's manager once onboarding completes (`/me/sendMail`). |
| `User.Invite.All` | Guest onboarding — `POST /invitations` instead of creating a cloud account. |
| `TeamMember.ReadWrite.All` | Add the new/cloned user to Teams selected in onboarding's Access step or a department template. |

SharePoint site access (the Access step's site grants) does not need a Graph
scope at all — `SiteAccessService` uses the operator's own delegated SharePoint
context via PnPJS (`sp.web.ensureUser` + associated group membership) against
the target site, the same way the rest of this web part talks to SharePoint.

Previously requested and removed, then reinstated once the corresponding
feature shipped: `TeamMember.ReadWrite.All`, `User.Invite.All`, `Mail.Send`.
Still not requested because nothing in the app calls it: `Sites.Read.All`,
`Mail.ReadBasic.All` (the latter backed a mailbox-readiness poll that was
removed — see below).

### Residual risk (stated once)

Delegated grants attach to the tenant-wide SharePoint Online Client
Extensibility principal, so **every** deployed SPFx solution in the tenant can
use these scopes. Mitigation is governance: restrict app-catalog deployment
rights (domain-isolated web parts are retired and not an option). Because
effective access still requires the caller's own Entra role, the practical
exposure is limited to what tenant admins can already do.

## Consequences accepted by this architecture

- **Approval workflow is advisory.** Enforced by the UI and the job state
  machine, not by a server. An Entra admin could bypass it — as they could in
  the admin centers.
- **No unattended execution.** Scheduled jobs run when an authorized operator
  has the app open; the dashboard surfaces due jobs with a "Run now" banner.
  Nothing executes without an operator session.
- **Audit is best-effort.** Audit entries are written client-side to
  `UPC_AuditLog`. Item-level permissions prevent edits by ordinary members,
  but a site admin could tamper. Accepted for v1.
- **Credentials are ephemeral.** Temporary passwords and TAPs are generated /
  retrieved client-side, displayed once in a copy-once dialog, and never
  written to any list, log or state.
- **Mailbox provisioning is not awaited.** Creating a user does not create an
  Exchange Online mailbox — that happens asynchronously, on Exchange's own
  timeline, only after license assignment succeeds. The engine assigns
  licenses and moves on; it does not poll for or block on mailbox readiness.
  Anything that would depend on the mailbox existing (e.g. sending mail to
  the new user immediately after the job completes) needs to account for
  that delay itself.

## Known limitations — operations Graph cannot perform

Each of these generates a `UPC_Tasks` record on the job instead of failing
silently. Future path: one Azure Automation runbook with Managed Identity —
added later without changing the web part.

| Operation | Why not Graph |
| --- | --- |
| Convert mailbox to shared | `Set-Mailbox -Type Shared` (EXO PowerShell only) |
| Mailbox delegation (Full Access / Send As / Send on Behalf) | `Add-MailboxPermission` / `Add-RecipientPermission` only |
| Admin-side email forwarding | `Set-Mailbox -ForwardingAddress` |
| Hide from GAL (mailbox users) | `Set-Mailbox -HiddenFromAddressListsEnabled`; Graph `showInAddressList` is deprecated/unreliable |
| Distribution list membership writes | Classic Exchange DLs are not writable via Graph |
| Litigation hold / retention | EXO / Purview |
| OneDrive archival / transfer to manager | SPO admin / Purview territory |
| Auto-reply on departed mailbox (admin-set) | Requires mailbox access; bundled with conversion |
| Hardware / asset provisioning | Not an M365 concept |
| Third-party app provisioning | No on-demand Graph API; group-driven Entra provisioning instead |
| License pricing | Graph exposes no cost data → manual `UPC_LicenseCostTable` |
| Unattended scheduled execution | No server in this architecture |

## Deployment

1. **Provision the lists** on the host site — either path, both idempotent:
   - *Property pane (no PowerShell):* after adding the web part, edit it and
     click **Provision UPC lists** in the property pane. Requires site owner
     (Manage Lists) rights on the host site. This also seeds the default
     `UPC_Settings` row (Title='app') with the factory settings JSON and the
     six `UPC_Roles` rows (ITAdmin, HRAdmin, DepartmentManager, ServiceDesk,
     Auditor, ReadOnly) with the canonical permission sets pre-filled — the
     admin only needs to paste the Entra security group object id into
     `MemberGroupId` for each role.
   - *Scripted:*
     ```powershell
     ./lists.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>"
     ```
   Then fill `UPC_Roles` `MemberGroupId` (Entra security group ids) and
   `UPC_LicenseCostTable` (Title = skuPartNumber, e.g. `SPE_E5`).

2. **Upload the package.** Build with `npm run build` and upload
   `sharepoint/solution/user-provisioning-center.sppkg` to the tenant app
   catalog.

3. **Approve the Graph permission requests.** SharePoint admin center →
   Advanced → **API access**: approve the eight pending Microsoft Graph scopes
   listed above. This is a Global Administrator action. If you're upgrading
   from a version that requested a different scope set, scopes no longer
   requested stay listed as previously approved until an admin revokes them —
   that's a manual cleanup step, this app doesn't (and can't) revoke its own
   consent.

4. **Governance note.** These delegated scopes become usable by any SPFx
   solution in the tenant (see residual risk). Restrict who may deploy to the
   app catalog.

5. **Add the web part** ("User Provisioning Center") to a page on the host
   site. No other infrastructure is required — there is no server component.

## Operator prerequisites

Operators need (a) membership in an Entra group referenced by `UPC_Roles` so
the UI offers actions, and (b) the actual Entra directory roles from the
matrix above so Graph accepts the writes. The preflight banner on the
dashboard tells each operator exactly what they are missing.
