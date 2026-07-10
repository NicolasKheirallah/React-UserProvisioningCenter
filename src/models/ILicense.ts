/** A subscribed SKU as read from Graph /subscribedSkus. */
export interface ISubscribedSku {
  skuId: string;
  skuPartNumber: string;
  /** Human-friendly product name (from a known-SKU lookup, falls back to skuPartNumber). */
  displayName: string;
  enabledUnits: number;
  consumedUnits: number;
  /** enabledUnits - consumedUnits, floored at 0. */
  availableUnits: number;
  capabilityStatus: string;
  /** True when the SKU includes an Exchange Online service plan. */
  includesExchange: boolean;
}

/** Manually maintained pricing row from UPC_LicenseCostTable (Graph has no pricing). */
export interface ILicenseCost {
  skuPartNumber: string;
  monthlyCost: number;
  currency: string;
}

/** SKU joined with its cost-table row for the license picker. */
export interface ILicenseOption extends ISubscribedSku {
  monthlyCost: number | null;
  currency: string | null;
}
