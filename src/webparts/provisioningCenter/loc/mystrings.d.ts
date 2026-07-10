declare interface IProvisioningCenterWebPartStrings {
  PropertyPaneDescription: string;
  ProvisionGroupName: string;
  ProvisionIntro: string;
  ProvisionButtonLabel: string;
  ProvisionStatusIdle: string;
  ProvisionStatusRunning: string;
  ProvisionStatusDone: string;
  ProvisionStatusError: string;
  ProvisionCreatedLabel: string;
  ProvisionExistingLabel: string;
  ProvisionFieldsLabel: string;
  ProvisionItemsLabel: string;
}

declare module 'ProvisioningCenterWebPartStrings' {
  const strings: IProvisioningCenterWebPartStrings;
  export = strings;
}
