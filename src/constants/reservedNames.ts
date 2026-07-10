/**
 * Baseline reserved UPN local parts. Additional entries can be supplied
 * from configuration (UPC_DepartmentTemplates admin settings) and are merged
 * at resolution time by the naming policy engine.
 */
export const DEFAULT_RESERVED_NAMES: readonly string[] = [
  'admin',
  'administrator',
  'security',
  'abuse',
  'postmaster',
  'noreply'
];
