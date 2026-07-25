/**
 * Returns the dedicated portal mount node for modal dialogs.
 * Falls back to document.body if the app root has not rendered yet.
 */
export function getDialogMountNode(): HTMLElement | undefined {
  if (typeof document === 'undefined') {
    return undefined;
  }
  return (document.getElementById('upc-dialog-root') as HTMLElement | null) ?? document.body;
}
