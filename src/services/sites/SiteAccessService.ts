import { spfi, SPFx, type SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/site-users/web';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

export type SiteAccessRole = 'visitor' | 'member' | 'owner';

/**
 * Grants a user access to an arbitrary SharePoint site (not necessarily the
 * host site the web part runs on) by adding them to that site's associated
 * Visitors/Members/Owners group — the same mechanism SharePoint's own
 * "Share" UI uses under the hood. This deliberately goes through PnPJS
 * against the target site directly, using the operator's own SharePoint
 * permissions there via SPFx's delegated context — no Microsoft Graph scope
 * is requested or required for this, unlike every other access-grant step.
 */
export class SiteAccessService {
  private readonly _context: WebPartContext;

  public constructor(context: WebPartContext) {
    this._context = context;
  }

  /**
   * Idempotent: SharePoint treats adding an existing group member as a
   * no-op, so this is safe to call again on a resumed job.
   */
  public async grantAccess(siteUrl: string, upn: string, role: SiteAccessRole): Promise<void> {
    const sp: SPFI = spfi(siteUrl).using(SPFx(this._context));
    const ensured: { LoginName: string } = await sp.web.ensureUser(upn);
    const group =
      role === 'owner'
        ? sp.web.associatedOwnerGroup
        : role === 'member'
          ? sp.web.associatedMemberGroup
          : sp.web.associatedVisitorGroup;
    await group.users.add(ensured.LoginName);
  }
}
