<#
.SYNOPSIS
  Provisions the UPC_* SharePoint lists for the User Provisioning Center web part.

.DESCRIPTION
  Creates the eleven lists defined by the solution's data contract (spec
  Section 4) on the target site. Idempotent: existing lists/fields are left
  in place.

  Requires the PnP.PowerShell module and site owner rights on the target site.

.EXAMPLE
  ./lists.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/it-provisioning"
#>
param(
  [Parameter(Mandatory = $true)]
  [string]$SiteUrl,

  # Entra app registration to use for the PnP connection (PnP interactive login).
  [Parameter(Mandatory = $false)]
  [string]$ClientId
)

$ErrorActionPreference = 'Stop'

if ($ClientId) {
  Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
} else {
  Connect-PnPOnline -Url $SiteUrl -Interactive
}

function Ensure-List {
  param([string]$Title)
  $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
  if ($null -eq $list) {
    Write-Host "Creating list $Title"
    $list = New-PnPList -Title $Title -Template GenericList -OnQuickLaunch:$false
  } else {
    Write-Host "List $Title already exists"
  }
  return $list
}

function Ensure-Field {
  param(
    [string]$List,
    [string]$InternalName,
    [string]$DisplayName,
    [string]$Type,           # Text | Note | Boolean | Number | Currency | DateTime | User | Choice
    [string[]]$Choices = @(),
    [switch]$Required
  )
  $existing = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
  if ($null -ne $existing) { return }
  Write-Host "  Adding field $List.$InternalName ($Type)"
  if ($Type -eq 'Choice') {
    Add-PnPField -List $List -InternalName $InternalName -DisplayName $DisplayName -Type Choice `
      -Choices $Choices -AddToDefaultView -Required:$Required | Out-Null
  } else {
    Add-PnPField -List $List -InternalName $InternalName -DisplayName $DisplayName -Type $Type `
      -AddToDefaultView -Required:$Required | Out-Null
  }
}

# Idempotent seed-row helper — mirrors ListProvisioningService._ensureSeedItems
# (matched by Title so re-running the script never duplicates a row; an
# existing row, even one an admin has since edited, is left untouched).
function Ensure-Item {
  param(
    [string]$List,
    [string]$Title,
    [hashtable]$Fields
  )
  $existing = Get-PnPListItem -List $List -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$Title</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
  if ($null -ne $existing -and $existing.Count -gt 0) {
    Write-Host "  Item $List/$Title already exists"
    return
  }
  Write-Host "  Adding item $List/$Title"
  $values = @{ Title = $Title } + $Fields
  Add-PnPListItem -List $List -Values $values | Out-Null
}

# ---------------------------------------------------------------- UPC_DepartmentTemplates
Ensure-List 'UPC_DepartmentTemplates' | Out-Null
Ensure-Field UPC_DepartmentTemplates TemplateJson 'Template JSON' Note
Ensure-Field UPC_DepartmentTemplates IsActive 'Is Active' Boolean
Ensure-Field UPC_DepartmentTemplates Version 'Version' Number

# ---------------------------------------------------------------- UPC_ProvisioningJobs
Ensure-List 'UPC_ProvisioningJobs' | Out-Null
Ensure-Field UPC_ProvisioningJobs JobType 'Job Type' Choice @('Onboard','Offboard','Transfer','Clone','Bulk') -Required
Ensure-Field UPC_ProvisioningJobs Status 'Status' Choice @('Draft','PendingApproval','Approved','Scheduled','Running','PartiallyFailed','Failed','Completed','Cancelled') -Required
Ensure-Field UPC_ProvisioningJobs PayloadJson 'Payload JSON' Note
Ensure-Field UPC_ProvisioningJobs StepsJson 'Steps JSON' Note
Ensure-Field UPC_ProvisioningJobs ScheduledFor 'Scheduled For' DateTime
Ensure-Field UPC_ProvisioningJobs RequestedBy 'Requested By' User
Ensure-Field UPC_ProvisioningJobs ApprovedBy 'Approved By' User
Ensure-Field UPC_ProvisioningJobs CorrelationId 'Correlation Id' Text
Ensure-Field UPC_ProvisioningJobs TargetUserId 'Target User Id' Text

# ---------------------------------------------------------------- UPC_AuditLog
Ensure-List 'UPC_AuditLog' | Out-Null
Ensure-Field UPC_AuditLog JobId 'Job Id' Text
Ensure-Field UPC_AuditLog Actor 'Actor' Text
Ensure-Field UPC_AuditLog Action 'Action' Text
Ensure-Field UPC_AuditLog TargetUser 'Target User' Text
Ensure-Field UPC_AuditLog GraphEndpoint 'Graph Endpoint' Text
Ensure-Field UPC_AuditLog RequestSummary 'Request Summary' Note
Ensure-Field UPC_AuditLog ResponseCode 'Response Code' Number
Ensure-Field UPC_AuditLog DurationMs 'Duration (ms)' Number
Ensure-Field UPC_AuditLog Result 'Result' Choice @('Success','Failure','Skipped')
Ensure-Field UPC_AuditLog CorrelationId 'Correlation Id' Text
Ensure-Field UPC_AuditLog TimestampUtc 'Timestamp (UTC)' DateTime

# Item-level security: members create entries but cannot edit anything afterwards.
# WriteSecurity 2 = users may only edit their own items; combined with versioning
# this is the closest list-level approximation of create-only. A site admin can
# still tamper — accepted, documented residual risk (spec Section 1).
Set-PnPList -Identity 'UPC_AuditLog' -ReadSecurity 1 -WriteSecurity 2 -EnableVersioning | Out-Null

# ---------------------------------------------------------------- UPC_ApplicationCatalog
Ensure-List 'UPC_ApplicationCatalog' | Out-Null
Ensure-Field UPC_ApplicationCatalog Owner 'Owner' User
Ensure-Field UPC_ApplicationCatalog ProvisioningType 'Provisioning Type' Choice @('Manual','GroupBased')
Ensure-Field UPC_ApplicationCatalog TargetGroupId 'Target Group Id' Text
Ensure-Field UPC_ApplicationCatalog ApprovalRequired 'Approval Required' Boolean
Ensure-Field UPC_ApplicationCatalog Instructions 'Instructions' Note
Ensure-Field UPC_ApplicationCatalog IsActive 'Is Active' Boolean

# ---------------------------------------------------------------- UPC_Roles
Ensure-List 'UPC_Roles' | Out-Null
Ensure-Field UPC_Roles MemberGroupId 'Member Group Id (Entra)' Text -Required
Ensure-Field UPC_Roles PermissionsJson 'Permissions JSON' Note

# ---------------------------------------------------------------- UPC_TeamsCatalog
Ensure-List 'UPC_TeamsCatalog' | Out-Null
Ensure-Field UPC_TeamsCatalog TeamId 'Team Id' Text -Required
Ensure-Field UPC_TeamsCatalog Category 'Category' Text
Ensure-Field UPC_TeamsCatalog DefaultRole 'Default Role' Choice @('member','owner')

# ---------------------------------------------------------------- UPC_SiteCatalog
Ensure-List 'UPC_SiteCatalog' | Out-Null
Ensure-Field UPC_SiteCatalog SiteUrl 'Site Url' Text -Required
Ensure-Field UPC_SiteCatalog BusinessOwner 'Business Owner' User
Ensure-Field UPC_SiteCatalog Category 'Category' Text

# ---------------------------------------------------------------- UPC_LicenseCostTable
# Title = skuPartNumber. Manually maintained — Graph exposes no pricing.
Ensure-List 'UPC_LicenseCostTable' | Out-Null
Ensure-Field UPC_LicenseCostTable MonthlyCost 'Monthly Cost' Currency
Ensure-Field UPC_LicenseCostTable Currency 'Currency' Text

# ---------------------------------------------------------------- UPC_Tasks
# Landing place for everything Graph cannot do (spec Section 10).
Ensure-List 'UPC_Tasks' | Out-Null
Ensure-Field UPC_Tasks JobId 'Job Id' Text
Ensure-Field UPC_Tasks TaskType 'Task Type' Choice @('MailboxConvertShared','MailboxDelegation','MailForwarding','HideFromGal','DistributionList','LitigationHold','OneDriveHandling','AutoReply','Hardware','ThirdPartyApp','OnPremAdAccount','GroupAssignment','TeamAssignment','SharePointAccess','ApplicationAssignment','AccessReview','Other')
Ensure-Field UPC_Tasks Instructions 'Instructions' Note
Ensure-Field UPC_Tasks AssignedTo 'Assigned To' User
Ensure-Field UPC_Tasks Status 'Status' Choice @('Open','Done')
Ensure-Field UPC_Tasks CompletedBy 'Completed By' Text
Ensure-Field UPC_Tasks CompletedUtc 'Completed (UTC)' DateTime

# ---------------------------------------------------------------- UPC_Settings
# Tenant-shared app configuration; single row with Title = 'app', edited from
# the in-app Settings tab (manageSettings permission).
Ensure-List 'UPC_Settings' | Out-Null
Ensure-Field UPC_Settings SettingsJson 'Settings JSON' Note

# ---------------------------------------------------------------- UPC_ApprovalDelegations
# Lets an approver delegate approveJobs to someone else for a time window
# (manageDelegations permission). Title holds the delegating operator's UPN.
Ensure-List 'UPC_ApprovalDelegations' | Out-Null
Ensure-Field UPC_ApprovalDelegations DelegateUpn 'Delegate UPN' Text -Required
Ensure-Field UPC_ApprovalDelegations StartUtc 'Start (UTC)' DateTime
Ensure-Field UPC_ApprovalDelegations EndUtc 'End (UTC)' DateTime
Ensure-Field UPC_ApprovalDelegations Reason 'Reason' Text
Ensure-Field UPC_ApprovalDelegations IsActive 'Is Active' Boolean

# ---------------------------------------------------------------- Seed data
# Matches src/services/provisioning/listSeedItems.ts (the in-app "Provision
# Lists" button's seed step) so both provisioning paths leave the site in
# the same state: one editable UPC_Settings row, and six UPC_Roles rows with
# their canonical permission sets pre-filled (MemberGroupId left blank for
# the admin to fill in with each role's Entra security group object id).
Write-Host ''
Write-Host 'Seeding default list items...'

Ensure-Item 'UPC_Settings' 'app' @{
  SettingsJson = '{"requireApproval":true,"bulkRowLimit":100,"jobsRefreshSeconds":30}'
}

Ensure-Item 'UPC_Roles' 'ITAdmin' @{
  MemberGroupId   = ''
  PermissionsJson = '["createJobs","approveJobs","runJobs","retrySteps","skipSteps","cancelJobs","manageTemplates","viewAudit","manageTasks","manageSettings"]'
}
Ensure-Item 'UPC_Roles' 'HRAdmin' @{
  MemberGroupId   = ''
  PermissionsJson = '["createJobs","approveJobs","viewAudit"]'
}
Ensure-Item 'UPC_Roles' 'DepartmentManager' @{
  MemberGroupId   = ''
  PermissionsJson = '["createJobs"]'
}
Ensure-Item 'UPC_Roles' 'ServiceDesk' @{
  MemberGroupId   = ''
  PermissionsJson = '["createJobs","runJobs","retrySteps","manageTasks"]'
}
Ensure-Item 'UPC_Roles' 'Auditor' @{
  MemberGroupId   = ''
  PermissionsJson = '["viewAudit"]'
}
Ensure-Item 'UPC_Roles' 'ReadOnly' @{
  MemberGroupId   = ''
  PermissionsJson = '[]'
}

Write-Host ''
Write-Host 'All UPC_* lists are provisioned and seeded.' -ForegroundColor Green
Write-Host 'Next: open the Roles tab (or edit UPC_Roles directly) to set each role''s Entra security group id, and fill UPC_LicenseCostTable.'
