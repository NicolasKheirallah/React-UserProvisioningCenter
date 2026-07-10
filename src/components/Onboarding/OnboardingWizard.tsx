import * as React from 'react';
import { Button, Subtitle1, makeStyles, tokens } from '@fluentui/react-components';
import * as strings from 'UpcStrings';
import { isDraftDirty, useWizard } from '../../contexts/WizardContext';
import { useServices } from '../../contexts/ServicesContext';
import { ConfirmDialog } from '../Shared/ConfirmDialog';
import { AppErrorBoundary } from '../Shared/AppErrorBoundary';
import { WizardStepper } from './WizardStepper';
import { PersonalStep } from './steps/PersonalStep';
import { EmploymentStep } from './steps/EmploymentStep';
import { IdentityStep } from './steps/IdentityStep';
import { AccountSettingsStep } from './steps/AccountSettingsStep';
import { LicensesStep } from './steps/LicensesStep';
import { AccessStep } from './steps/AccessStep';
import { ReviewStep } from './steps/ReviewStep';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalL
  },
  titleRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    columnGap: tokens.spacingHorizontalM
  },
  body: {
    display: 'flex',
    columnGap: tokens.spacingHorizontalXXXL,
    alignItems: 'flex-start',
    // Narrow section/Teams mobile: stepper collapses above the content.
    '@container (max-width: 560px)': {
      flexDirection: 'column',
      rowGap: tokens.spacingVerticalL
    }
  },
  content: {
    flexGrow: 1,
    minWidth: 0,
    width: '100%'
  }
});

export interface IOnboardingWizardProps {
  onSubmitted: () => void;
}

// Static localized labels — computed once rather than reallocated on every
// render (this component re-renders on every wizard-draft keystroke).
const STEP_LABELS: string[] = [
  strings.WizardStepPersonal,
  strings.WizardStepEmployment,
  strings.WizardStepIdentity,
  strings.WizardStepAccount,
  strings.WizardStepLicenses,
  strings.WizardStepAccess,
  strings.WizardStepReview
];

/**
 * Personal, Employment, Identity, Account, Licenses, Access, Review. Draft
 * state lives in the app-level WizardProvider so it survives tab switches;
 * "Start over" is the explicit way to discard it.
 */
export const OnboardingWizard: React.FC<IOnboardingWizardProps> = ({ onSubmitted }) => {
  const styles = useStyles();
  const { state, dispatch } = useWizard();
  const services = useServices();
  const [confirmDiscard, setConfirmDiscard] = React.useState<boolean>(false);

  const onStepClick = React.useCallback(
    (step: number) => dispatch({ type: 'goto', step }),
    [dispatch]
  );
  const onStepRenderError = React.useCallback(
    (error: Error, info: React.ErrorInfo) =>
      services.telemetry.trackError(error, {
        componentStack: info.componentStack,
        scope: 'render',
        wizardStep: state.step
      }),
    [services, state.step]
  );

  return (
    <div className={styles.root}>
      <div className={styles.titleRow}>
        <Subtitle1 as="h3" block>
          {strings.WizardTitle}
        </Subtitle1>
        {isDraftDirty(state) ? (
          <Button appearance="subtle" size="small" onClick={() => setConfirmDiscard(true)}>
            {strings.StartOverLabel}
          </Button>
        ) : undefined}
      </div>
      <div className={styles.body}>
        <WizardStepper labels={STEP_LABELS} current={state.step} onStepClick={onStepClick} />
        <div className={styles.content}>
          {/* Keyed on the step index: a crash in one step's render no longer
              blanks the whole wizard (stepper/title survive), and navigating
              away from the broken step — even without an explicit Retry —
              remounts fresh and recovers. */}
          <AppErrorBoundary key={state.step} onError={onStepRenderError}>
            {state.step === 0 ? <PersonalStep /> : undefined}
            {state.step === 1 ? <EmploymentStep /> : undefined}
            {state.step === 2 ? <IdentityStep /> : undefined}
            {state.step === 3 ? <AccountSettingsStep /> : undefined}
            {state.step === 4 ? <LicensesStep /> : undefined}
            {state.step === 5 ? <AccessStep /> : undefined}
            {state.step === 6 ? <ReviewStep onSubmitted={onSubmitted} /> : undefined}
          </AppErrorBoundary>
        </div>
      </div>
      <ConfirmDialog
        open={confirmDiscard}
        title={strings.DiscardDraftTitle}
        message={strings.DiscardDraftBody}
        confirmLabel={strings.DiscardDraftConfirm}
        cancelLabel={strings.KeepEditingLabel}
        onConfirm={() => {
          dispatch({ type: 'reset' });
          setConfirmDiscard(false);
        }}
        onCancel={() => setConfirmDiscard(false)}
      />
    </div>
  );
};
