/** Substitutes {0}, {1}, … placeholders in a localized string. */
export function formatString(template: string, ...args: string[]): string {
  return template.replace(/\{(\d+)\}/g, (match, index) => args[Number(index)] ?? match);
}
