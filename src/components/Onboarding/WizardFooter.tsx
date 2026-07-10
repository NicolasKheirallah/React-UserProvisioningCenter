import * as React from 'react';
import { Button, Spinner, makeStyles, tokens } from '@fluentui/react-components';
import * as strings from 'UpcStrings';

const useStyles = makeStyles({
  // The divider spans the whole content column (admin-center pattern) so it
  // lines up for both narrow forms and wide table steps.
  footer: {
    display: 'flex',
    columnGap: tokens.spacingHorizontalS,
    paddingTop: tokens.spacingVerticalL,
    marginTop: tokens.spacingVerticalL,
    borderTopWidth: tokens.strokeWidthThin,
    borderTopStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2
  },
  button: {
    minWidth: '96px'
  }
});

export interface IWizardFooterProps {
  onBack?: () => void;
  /** Omit when the hosting form submits via type="submit". */
  onNext?: () => void;
  nextDisabled?: boolean;
  nextLoading?: boolean;
  nextLabel?: string;
  nextIsSubmit?: boolean;
}

/**
 * Admin-center wizard footer: divider above, fixed-min-width Back/Next pair.
 * While busy the primary button keeps its label and shows a spinner in the
 * icon slot, so its width and meaning stay stable.
 */
export const WizardFooter: React.FC<IWizardFooterProps> = (props) => {
  const styles = useStyles();
  return (
    <div className={styles.footer}>
      {props.onBack ? (
        <Button appearance="secondary" className={styles.button} onClick={props.onBack}>
          {strings.BackLabel}
        </Button>
      ) : undefined}
      <Button
        appearance="primary"
        className={styles.button}
        type={props.nextIsSubmit ? 'submit' : 'button'}
        onClick={props.onNext}
        disabled={props.nextDisabled || props.nextLoading}
        icon={
          props.nextLoading ? (
            <Spinner size="tiny" aria-label={strings.LoadingLabel} />
          ) : undefined
        }
      >
        {props.nextLabel ?? strings.NextLabel}
      </Button>
    </div>
  );
};
