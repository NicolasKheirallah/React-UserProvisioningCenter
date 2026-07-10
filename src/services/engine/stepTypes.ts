import type { GraphService } from '../graph/GraphService';
import type { SharePointDataService } from '../sharePointData/SharePointDataService';
import type { AuditService } from '../audit/AuditService';
import type { NamingPolicyService } from '../namingPolicy/NamingPolicyService';
import type { UserService } from '../users/UserService';
import type { SiteAccessService } from '../sites/SiteAccessService';
import type { CredentialMode, IAppSettings, IProvisioningJob } from '../../models';

/**
 * Secrets exist only in this in-memory object for the lifetime of a run.
 * They are never serialized to StepsJson, PayloadJson, audit or anywhere else.
 */
export interface IJobSecrets {
  temporaryPassword?: string;
  temporaryAccessPass?: string;
}

export interface ICredentialPresentation {
  kind: CredentialMode;
  value: string;
  userPrincipalName: string;
}

export interface IStepContext {
  graph: GraphService;
  data: SharePointDataService;
  audit: AuditService;
  naming: NamingPolicyService;
  users: UserService;
  siteAccess: SiteAccessService;
  /** Mutable during the run; targetUserId is set by create-user. */
  job: IProvisioningJob;
  secrets: IJobSecrets;
  signal?: AbortSignal;
  /** Tenant-shared app settings (read-only) so steps can read runtime config. */
  settings: IAppSettings;
  /**
   * Hands a one-time credential to the UI (copy-once Dialog addressed to the
   * operator). Resolves when the operator confirms they have secured it.
   */
  presentCredentials: (credential: ICredentialPresentation) => Promise<void>;
}

/** Thrown by steps; drives the retry policy and StepsJson lastError. */
export class StepFailure extends Error {
  public readonly graphCode: string;
  public readonly retryable: boolean;

  public constructor(message: string, graphCode: string, retryable: boolean) {
    super(message);
    Object.setPrototypeOf(this, StepFailure.prototype);
    this.name = 'StepFailure';
    this.graphCode = graphCode;
    this.retryable = retryable;
  }
}

export interface IWorkflowStepDefinition {
  id: string;
  skippable: boolean;
  maxAttempts: number;
  /**
   * When true a permanent failure does not block later steps — the job
   * continues and ends PartiallyFailed instead of Failed (Section 6).
   */
  continueOnFailure: boolean;
  run: (ctx: IStepContext) => Promise<void>;
}
