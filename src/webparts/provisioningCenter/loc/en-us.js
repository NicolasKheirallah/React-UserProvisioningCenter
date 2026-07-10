define([], function () {
  return {
    PropertyPaneDescription: 'User Provisioning Center configuration',
    ProvisionGroupName: 'Data store',
    ProvisionIntro:
      'Creates the UPC_* SharePoint lists on this site (same result as provisioning-assets/lists.ps1). Requires site owner rights. Safe to re-run: existing lists and columns are left untouched.',
    ProvisionButtonLabel: 'Provision UPC lists',
    ProvisionStatusIdle: 'Lists have not been checked yet.',
    ProvisionStatusRunning: 'Provisioning…',
    ProvisionStatusDone: 'Done.',
    ProvisionStatusError: 'Provisioning failed:',
    ProvisionCreatedLabel: 'lists created',
    ProvisionExistingLabel: 'already existed',
    ProvisionFieldsLabel: 'columns added',
    ProvisionItemsLabel: 'list items seeded'
  };
});
