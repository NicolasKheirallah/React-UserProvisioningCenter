import {
  LIST_APPLICATION_CATALOG,
  LIST_AUDIT_LOG,
  LIST_DEPARTMENT_TEMPLATES,
  LIST_LICENSE_COST_TABLE,
  LIST_PROVISIONING_JOBS,
  LIST_ROLES,
  LIST_SITE_CATALOG,
  LIST_SETTINGS,
  LIST_TASKS,
  LIST_TEAMS_CATALOG
} from '../../constants/listNames';

export type UpcFieldType =
  | 'Text'
  | 'Note'
  | 'Boolean'
  | 'Number'
  | 'Currency'
  | 'DateTime'
  | 'User'
  | 'Choice';

export interface IUpcFieldDefinition {
  /** Internal/static name — must match what the data services select. */
  name: string;
  displayName: string;
  type: UpcFieldType;
  required?: boolean;
  choices?: string[];
}

export interface IUpcListDefinition {
  title: string;
  fields: IUpcFieldDefinition[];
  /** UPC_AuditLog: restrict members to editing only their own items + versioning. */
  auditSecurity?: boolean;
}

/**
 * The nine UPC_* list schemas (spec Section 4). Single source for the
 * property-pane provisioning path; provisioning-assets/lists.ps1 is the
 * scripted equivalent — keep both in sync.
 */
export const UPC_LIST_DEFINITIONS: IUpcListDefinition[] = [
  {
    title: LIST_DEPARTMENT_TEMPLATES,
    fields: [
      { name: 'TemplateJson', displayName: 'Template JSON', type: 'Note' },
      { name: 'IsActive', displayName: 'Is Active', type: 'Boolean' },
      { name: 'Version', displayName: 'Version', type: 'Number' }
    ]
  },
  {
    title: LIST_PROVISIONING_JOBS,
    fields: [
      {
        name: 'JobType',
        displayName: 'Job Type',
        type: 'Choice',
        required: true,
        choices: ['Onboard', 'Offboard', 'Transfer', 'Clone', 'Bulk']
      },
      {
        name: 'Status',
        displayName: 'Status',
        type: 'Choice',
        required: true,
        choices: [
          'Draft',
          'PendingApproval',
          'Approved',
          'Scheduled',
          'Running',
          'PartiallyFailed',
          'Failed',
          'Completed',
          'Cancelled'
        ]
      },
      { name: 'PayloadJson', displayName: 'Payload JSON', type: 'Note' },
      { name: 'StepsJson', displayName: 'Steps JSON', type: 'Note' },
      { name: 'ScheduledFor', displayName: 'Scheduled For', type: 'DateTime' },
      { name: 'RequestedBy', displayName: 'Requested By', type: 'User' },
      { name: 'ApprovedBy', displayName: 'Approved By', type: 'User' },
      { name: 'CorrelationId', displayName: 'Correlation Id', type: 'Text' },
      { name: 'TargetUserId', displayName: 'Target User Id', type: 'Text' }
    ]
  },
  {
    title: LIST_AUDIT_LOG,
    auditSecurity: true,
    fields: [
      { name: 'JobId', displayName: 'Job Id', type: 'Text' },
      { name: 'Actor', displayName: 'Actor', type: 'Text' },
      { name: 'Action', displayName: 'Action', type: 'Text' },
      { name: 'TargetUser', displayName: 'Target User', type: 'Text' },
      { name: 'GraphEndpoint', displayName: 'Graph Endpoint', type: 'Text' },
      { name: 'RequestSummary', displayName: 'Request Summary', type: 'Note' },
      { name: 'ResponseCode', displayName: 'Response Code', type: 'Number' },
      { name: 'DurationMs', displayName: 'Duration (ms)', type: 'Number' },
      {
        name: 'Result',
        displayName: 'Result',
        type: 'Choice',
        choices: ['Success', 'Failure', 'Skipped']
      },
      { name: 'CorrelationId', displayName: 'Correlation Id', type: 'Text' },
      { name: 'TimestampUtc', displayName: 'Timestamp (UTC)', type: 'DateTime' }
    ]
  },
  {
    title: LIST_APPLICATION_CATALOG,
    fields: [
      { name: 'Owner', displayName: 'Owner', type: 'User' },
      {
        name: 'ProvisioningType',
        displayName: 'Provisioning Type',
        type: 'Choice',
        choices: ['Manual', 'GroupBased']
      },
      { name: 'TargetGroupId', displayName: 'Target Group Id', type: 'Text' },
      { name: 'ApprovalRequired', displayName: 'Approval Required', type: 'Boolean' },
      { name: 'Instructions', displayName: 'Instructions', type: 'Note' },
      { name: 'IsActive', displayName: 'Is Active', type: 'Boolean' }
    ]
  },
  {
    title: LIST_ROLES,
    fields: [
      {
        name: 'MemberGroupId',
        displayName: 'Member Group Id (Entra)',
        type: 'Text',
        required: true
      },
      { name: 'PermissionsJson', displayName: 'Permissions JSON', type: 'Note' }
    ]
  },
  {
    title: LIST_TEAMS_CATALOG,
    fields: [
      { name: 'TeamId', displayName: 'Team Id', type: 'Text', required: true },
      { name: 'Category', displayName: 'Category', type: 'Text' },
      {
        name: 'DefaultRole',
        displayName: 'Default Role',
        type: 'Choice',
        choices: ['member', 'owner']
      }
    ]
  },
  {
    title: LIST_SITE_CATALOG,
    fields: [
      { name: 'SiteUrl', displayName: 'Site Url', type: 'Text', required: true },
      { name: 'BusinessOwner', displayName: 'Business Owner', type: 'User' },
      { name: 'Category', displayName: 'Category', type: 'Text' }
    ]
  },
  {
    // Title = skuPartNumber. Manually maintained — Graph exposes no pricing.
    title: LIST_LICENSE_COST_TABLE,
    fields: [
      { name: 'MonthlyCost', displayName: 'Monthly Cost', type: 'Currency' },
      { name: 'Currency', displayName: 'Currency', type: 'Text' }
    ]
  },
  {
    // Landing place for everything Graph cannot do (spec Section 10).
    title: LIST_TASKS,
    fields: [
      { name: 'JobId', displayName: 'Job Id', type: 'Text' },
      {
        name: 'TaskType',
        displayName: 'Task Type',
        type: 'Choice',
        choices: [
          'MailboxConvertShared',
          'MailboxDelegation',
          'MailForwarding',
          'HideFromGal',
          'DistributionList',
          'LitigationHold',
          'OneDriveHandling',
          'AutoReply',
          'Hardware',
          'ThirdPartyApp',
          'OnPremAdAccount',
          'GroupAssignment',
          'TeamAssignment',
          'SharePointAccess',
          'ApplicationAssignment',
          'AccessReview',
          'Other'
        ]
      },
      { name: 'Instructions', displayName: 'Instructions', type: 'Note' },
      { name: 'AssignedTo', displayName: 'Assigned To', type: 'User' },
      { name: 'Status', displayName: 'Status', type: 'Choice', choices: ['Open', 'Done'] },
      { name: 'CompletedBy', displayName: 'Completed By', type: 'Text' },
      { name: 'CompletedUtc', displayName: 'Completed (UTC)', type: 'DateTime' }
    ]
  },
  {
    // Tenant-shared app configuration; single row with Title = 'app'.
    title: LIST_SETTINGS,
    fields: [{ name: 'SettingsJson', displayName: 'Settings JSON', type: 'Note' }]
  }
];
