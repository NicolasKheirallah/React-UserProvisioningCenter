export interface ISubscribedSku {
  skuId: string;
  skuPartNumber: string;
  displayName: string;
  enabledUnits: number;
  consumedUnits: number;
  availableUnits: number;
  capabilityStatus: string;
  includesExchange: boolean;
}

export interface ILicenseCost {
  skuPartNumber: string;
  monthlyCost: number;
  currency: string;
}

export interface ILicenseOption extends ISubscribedSku {
  monthlyCost: number | null;
  currency: string | null;
}
