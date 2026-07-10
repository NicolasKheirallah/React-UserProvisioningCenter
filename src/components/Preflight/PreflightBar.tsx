import * as React from 'react';
import {
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  makeStyles,
  tokens
} from '@fluentui/react-components';
import * as strings from 'UpcStrings';
import { usePreflight } from '../../hooks/useReferenceData';

const useStyles = makeStyles({
  list: {
    margin: 0,
    paddingLeft: tokens.spacingHorizontalXL
  }
});

/**
 * Permission preflight surface (spec Section 1): shown at app load when the
 * signed-in operator lacks Entra roles for some capabilities. Purely
 * informational — enforcement happens in Entra, not here.
 */
export const PreflightBar: React.FC = () => {
  const styles = useStyles();
  const preflight = usePreflight();

  if (preflight.isLoading) {
    return null;
  }
  if (preflight.error) {
    return (
      <MessageBar intent="error" layout="multiline">
        <MessageBarBody>
          <MessageBarTitle>{strings.PreflightErrorTitle}</MessageBarTitle>
          {preflight.error instanceof Error ? preflight.error.message : ''}
        </MessageBarBody>
      </MessageBar>
    );
  }
  if (!preflight.data || preflight.data.missing.length === 0) {
    return null;
  }
  return (
    <MessageBar intent="warning" layout="multiline">
      <MessageBarBody>
        <MessageBarTitle>{strings.PreflightMissingTitle}</MessageBarTitle>
        {strings.PreflightMissingIntro}
        <ul className={styles.list}>
          {preflight.data.missing.map((check) => (
            <li key={check.capability}>
              <strong>{check.label}</strong> — {check.detail}
            </li>
          ))}
        </ul>
      </MessageBarBody>
    </MessageBar>
  );
};
