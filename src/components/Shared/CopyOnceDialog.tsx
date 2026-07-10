import * as React from 'react';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Text,
  makeStyles,
  tokens
} from '@fluentui/react-components';
import * as strings from 'UpcStrings';
import type { ICredentialPresentation } from '../../services/engine/stepTypes';

/** Clear the OS clipboard this long after a copy, so the secret doesn't sit
 *  there indefinitely for another app or a later paste to pick up. */
const CLIPBOARD_CLEAR_DELAY_MS: number = 45 * 1000;

const useStyles = makeStyles({
  secret: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase500,
    padding: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    textAlign: 'center',
    userSelect: 'all',
    letterSpacing: '0.05em'
  },
  body: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalM
  }
});

export interface ICopyOnceDialogProps {
  credential: ICredentialPresentation | null;
  /** Called once the operator confirms secure hand-off; the value is then gone. */
  onConfirm: () => void;
}

/**
 * Copy-once credential dialog (spec Section 6). The value exists only in
 * memory, is displayed exactly once and is never persisted. The dialog is
 * deliberately non-dismissable except through explicit confirmation.
 */
export const CopyOnceDialog: React.FC<ICopyOnceDialogProps> = ({ credential, onConfirm }) => {
  const styles = useStyles();
  const [copied, setCopied] = React.useState<boolean>(false);

  React.useEffect(() => {
    setCopied(false);
  }, [credential]);

  if (!credential) {
    return null;
  }

  // Best-effort: clear the clipboard after a delay, but only if it still
  // holds exactly what we copied — never clobber something copied since.
  // readText() needs the clipboard-read permission and may be denied or
  // unsupported in some host contexts (SharePoint/Teams iframes); silently
  // leaving the clipboard alone is the safe fallback either way.
  const scheduleClipboardClear = (value: string): void => {
    window.setTimeout(() => {
      if (!navigator.clipboard?.readText || !navigator.clipboard?.writeText) {
        return;
      }
      navigator.clipboard
        .readText()
        .then((current) => (current === value ? navigator.clipboard.writeText('') : undefined))
        .catch(() => undefined);
    }, CLIPBOARD_CLEAR_DELAY_MS);
  };

  const copy = async (): Promise<void> => {
    // Feature-detect: navigator.clipboard is undefined in non-secure contexts
    // (HTTP, some iframes). Falling back to a non-copy state is safer than
    // letting the operator believe they copied the credential when they didn't.
    if (!navigator.clipboard?.writeText) {
      setCopied(false);
      return;
    }
    try {
      await navigator.clipboard.writeText(credential.value);
      setCopied(true);
      scheduleClipboardClear(credential.value);
    } catch {
      setCopied(false);
    }
  };

  const bodyLabel: string =
    credential.kind === 'tap'
      ? strings.CredentialDialogBodyTap
      : strings.CredentialDialogBodyPassword;

  return (
    <Dialog modalType="alert" open={true}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{strings.CredentialDialogTitle}</DialogTitle>
          <DialogContent className={styles.body}>
            <Text>
              {bodyLabel} <strong>{credential.userPrincipalName}</strong>
            </Text>
            <div className={styles.secret}>{credential.value}</div>
            <Text size={200}>{strings.CredentialDialogWarning}</Text>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={copy}>
              {copied ? strings.CopiedLabel : strings.CopyLabel}
            </Button>
            <Button appearance="primary" onClick={onConfirm}>
              {strings.CredentialDialogConfirm}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
