import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx, type SPFI } from '@pnp/sp';
import { GraphService } from './graph/GraphService';
import { SharePointDataService } from './sharePointData/SharePointDataService';
import { AuditService } from './audit/AuditService';
import { NamingPolicyService } from './namingPolicy/NamingPolicyService';
import { UserService } from './users/UserService';
import { LicenseService } from './licenses/LicenseService';
import { RoleService } from './roles/RoleService';
import { PreflightService } from './preflight/PreflightService';
import { WorkflowEngine } from './engine/WorkflowEngine';
import { TelemetryService } from './telemetry/TelemetryService';
import { SiteAccessService } from './sites/SiteAccessService';

export interface IServices {
  graph: GraphService;
  data: SharePointDataService;
  audit: AuditService;
  naming: NamingPolicyService;
  users: UserService;
  licenses: LicenseService;
  roles: RoleService;
  preflight: PreflightService;
  engine: WorkflowEngine;
  telemetry: TelemetryService;
  siteAccess: SiteAccessService;
  operatorUpn: string;
  /**
   * SharePoint's cached profile-photo endpoint for an account — cheap enough
   * for type-ahead results, unlike per-user Graph photo requests.
   */
  photoUrl: (upn: string) => string;
}

/** Composition root, called once from the web part's onInit. */
export async function createServices(context: WebPartContext): Promise<IServices> {
  const client = await context.msGraphClientFactory.getClient('3');
  const graph: GraphService = new GraphService(client);
  const sp: SPFI = spfi().using(SPFx(context));
  const data: SharePointDataService = new SharePointDataService(sp);
  const telemetry: TelemetryService = new TelemetryService();
  const operatorUpn: string = context.pageContext.user.loginName;
  const audit: AuditService = new AuditService(data, operatorUpn);
  const naming: NamingPolicyService = new NamingPolicyService(graph);
  const users: UserService = new UserService(graph);
  const licenses: LicenseService = new LicenseService(graph, data);
  const roles: RoleService = new RoleService(graph, data);
  const preflight: PreflightService = new PreflightService(graph);
  const siteAccess: SiteAccessService = new SiteAccessService(context);
  const engine: WorkflowEngine = new WorkflowEngine({
    graph,
    data,
    audit,
    naming,
    users,
    siteAccess,
    telemetry
  });
  graph.setTelemetry(telemetry);
  preflight.setTelemetry(telemetry);
  audit.setTelemetry(telemetry);
  data.setTelemetry(telemetry);
  const webUrl: string = context.pageContext.web.absoluteUrl;
  const photoUrl = (upn: string): string =>
    `${webUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(upn)}`;
  return {
    graph,
    data,
    audit,
    naming,
    users,
    licenses,
    roles,
    preflight,
    engine,
    telemetry,
    siteAccess,
    operatorUpn,
    photoUrl
  };
}
