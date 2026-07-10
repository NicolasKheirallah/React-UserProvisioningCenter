/** Escape single quotes in OData $filter literals. */
export function escapeODataLiteral(value: string): string {
  return value.replace(/'/g, "''");
}