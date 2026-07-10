import { batchTyped, type GraphService } from '../graph/GraphService';
import type { TelemetryService } from '../telemetry/TelemetryService';
import type { CapabilityId, ICapabilityCheck, IPreflightResult } from '../../models';

/** Well-known Entra directory role template ids. */
const ROLE_TEMPLATE: Record<string, string> = {
  globalAdministrator: '62e90394-69f5-4237-9190-012177145e10',
  userAdministrator: 'fe930be7-5e62-47db-91af-98c3a49a38b1',
  licenseAdministrator: '4d6ac14f-3453-41d0-bef9-a3e0c569773a',
  groupsAdministrator: 'fdd7a751-b60b-444a-984c-02652fe8fa1c',
  teamsAdministrator: '69091246-20e8-4a56-aa4d-066075b2a7a8',
  authenticationAdministrator: 'c4e39bd9-1100-46d3-8c65-fb160da0071f',
  privilegedAuthenticationAdministrator: '7be44c8a-adaf-4e2a-84d6-ab2649e08a13',
  guestInviter: '95e79109-95c0-4d8e-aee3-d01accf2d47b'
};

interface ICapabilityRule {
  capability: CapabilityId;
  label: string;
  detail: string;
  /** Holding any of these directory roles grants the capability. */
  satisfiedBy: string[];
}

/**
 * Capability → minimum Entra role matrix (spec Section 1). Global Admin
 * satisfies everything. Group/team-owner based access cannot be probed
 * cheaply and is mentioned in the detail text instead.
 */
const CAPABILITY_RULES: ICapabilityRule[] = [
  {
    capability: 'createUsers',
    label: 'Create and update users',
    detail: 'Requires the User Administrator role.',
    satisfiedBy: [ROLE_TEMPLATE.userAdministrator]
  },
  {
    capability: 'assignLicenses',
    label: 'Assign and remove licenses',
    detail: 'Requires the License Administrator (or User Administrator) role.',
    satisfiedBy: [ROLE_TEMPLATE.licenseAdministrator, ROLE_TEMPLATE.userAdministrator]
  },
  {
    capability: 'groupWrites',
    label: 'Manage group membership',
    detail: 'Requires the Groups Administrator role, or ownership of the target groups.',
    satisfiedBy: [ROLE_TEMPLATE.groupsAdministrator]
  },
  {
    capability: 'teamsWrites',
    label: 'Manage Teams membership',
    detail: 'Requires the Teams Administrator role, or ownership of the target teams.',
    satisfiedBy: [ROLE_TEMPLATE.teamsAdministrator]
  },
  {
    capability: 'tapCreation',
    label: 'Create Temporary Access Passes',
    detail: 'Requires the Authentication Administrator role.',
    satisfiedBy: [
      ROLE_TEMPLATE.authenticationAdministrator,
      ROLE_TEMPLATE.privilegedAuthenticationAdministrator
    ]
  },
  {
    capability: 'guestInvites',
    label: 'Invite guests',
    detail: 'Requires the Guest Inviter or User Administrator role.',
    satisfiedBy: [ROLE_TEMPLATE.guestInviter, ROLE_TEMPLATE.userAdministrator]
  },
  {
    capability: 'revokeSessions',
    label: 'Revoke sign-in sessions (offboarding)',
    detail: 'Requires the User Administrator role.',
    satisfiedBy: [ROLE_TEMPLATE.userAdministrator]
  }
];

interface IDirectoryRolesResponse {
  value: { roleTemplateId?: string; displayName?: string }[];
}

/**
 * Permission preflight, run at app load (spec Section 1). Read probes verify
 * the delegated grant is consented; the operator's active directory roles
 * (via /me/memberOf) determine which admin capabilities will actually work,
 * because delegated effective access = app scope ∩ operator privileges.
 */
export class PreflightService {
  private readonly _graph: GraphService;
  private _telemetry: TelemetryService | undefined;

  public constructor(graph: GraphService) {
    this._graph = graph;
  }

  public setTelemetry(telemetry: TelemetryService): void {
    this._telemetry = telemetry;
  }

  public async run(signal?: AbortSignal): Promise<IPreflightResult> {
    const checks: ICapabilityCheck[] = [];
    const started: number = Date.now();

    try {
      const me: { userPrincipalName: string } = await this._graph.get<{
        userPrincipalName: string;
      }>('/me?$select=userPrincipalName', { signal });

      // Harmless read probes: prove consent for the read scopes we depend on.
      // batchTyped keeps the id/url map in one typed object so a renamed or
      // mistyped probe id fails to compile instead of silently reading undefined.
      const probes = await batchTyped(
        this._graph,
        {
          org: { method: 'GET', url: '/organization?$select=id' },
          skus: { method: 'GET', url: '/subscribedSkus' },
          users: { method: 'GET', url: '/users?$top=1&$select=id' }
        },
        { signal }
      );
      const probeOk = (id: keyof typeof probes): boolean => {
        const status: number = probes[id]?.status ?? 0;
        return status >= 200 && status < 300;
      };
      checks.push({
        capability: 'directoryRead',
        label: 'Read the directory',
        detail: 'Directory.Read.All / Organization.Read.All consent missing or denied.',
        ok: probeOk('org') && probeOk('users')
      });
      checks.push({
        capability: 'licenseRead',
        label: 'Read subscribed licenses',
        detail: 'Organization.Read.All consent missing or denied.',
        ok: probeOk('skus')
      });

      let templateIds: string[] = [];
      try {
        const roles: IDirectoryRolesResponse = await this._graph.get<IDirectoryRolesResponse>(
          '/me/memberOf/microsoft.graph.directoryRole?$select=displayName,roleTemplateId',
          { signal }
        );
        templateIds = (roles.value ?? [])
          .map((r) => r.roleTemplateId ?? '')
          .filter((id) => id.length > 0);
      } catch {
        // If we cannot enumerate roles, report the admin capabilities as unknown-missing.
        templateIds = [];
      }
      const isGlobalAdmin: boolean = templateIds.indexOf(ROLE_TEMPLATE.globalAdministrator) !== -1;

      for (const rule of CAPABILITY_RULES) {
        const ok: boolean =
          isGlobalAdmin || rule.satisfiedBy.some((id) => templateIds.indexOf(id) !== -1);
        checks.push({ capability: rule.capability, label: rule.label, detail: rule.detail, ok });
      }

      const result: IPreflightResult = {
        checks,
        missing: checks.filter((c) => !c.ok),
        directoryRoleTemplateIds: templateIds,
        operatorUpn: me.userPrincipalName
      };
      if (this._telemetry) {
        this._telemetry.trackEvent(
          'preflight.run',
          {
            operatorUpn: me.userPrincipalName,
            missingCount: result.missing.length,
            globalAdmin: isGlobalAdmin
          },
          'info'
        );
      }
      return result;
    } catch (err) {
      if (this._telemetry) {
        this._telemetry.trackError(err, { durationMs: Date.now() - started });
      }
      throw err;
    }
  }
}
