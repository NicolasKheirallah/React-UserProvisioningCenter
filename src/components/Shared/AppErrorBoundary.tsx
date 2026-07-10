import * as React from 'react';
import {
  Button,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  makeStyles,
  tokens
} from '@fluentui/react-components';
import * as strings from 'UpcStrings';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalM
  }
});

interface IFallbackProps {
  error: Error;
  onRetry: () => void;
}

/** Function component so Fluent's makeStyles (a Hook) is legal here. */
const ErrorFallback: React.FC<IFallbackProps> = ({ error, onRetry }) => {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <MessageBar intent="error" layout="multiline">
        <MessageBarBody>
          <MessageBarTitle>{strings.ErrorGenericTitle}</MessageBarTitle>
          {error.message || strings.ErrorGenericTitle}
        </MessageBarBody>
        <Button appearance="secondary" onClick={onRetry}>
          {strings.RetryLabel}
        </Button>
      </MessageBar>
    </div>
  );
};

interface IState {
  error: Error | undefined;
  /** Incremented on retry to force a child remount. */
  resetNonce: number;
}

/**
 * Top-level error boundary. A render-time exception anywhere in the web part
 * surfaces a localized fallback with a Retry affordance instead of a blank
 * web part. SPFx's own error UI is coarse; this keeps the user in the app.
 *
 * Error boundaries MUST be class components (getDerivedStateFromError /
 * componentDidCatch have no Hook equivalent). The fallback UI is delegated to
 * a function component so Fluent's makeStyles Hook is legal.
 */
interface IAppErrorBoundaryProps {
  children: React.ReactNode;
  /** Reported alongside the console log — wire to TelemetryService.trackError. */
  onError?: (error: Error, info: React.ErrorInfo) => void;
}

export class AppErrorBoundary extends React.Component<IAppErrorBoundaryProps, IState> {
  public state: IState = { error: undefined, resetNonce: 0 };

  public static getDerivedStateFromError(error: Error): Partial<IState> {
    return { error };
  }

  public componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error('[UPC] render error', error, info.componentStack);
    this.props.onError?.(error, info);
  }

  private _retry = (): void => {
    this.setState({ error: undefined, resetNonce: this.state.resetNonce + 1 });
  };

  public render(): React.ReactNode {
    if (this.state.error) {
      return <ErrorFallback error={this.state.error} onRetry={this._retry} />;
    }
    return <React.Fragment key={this.state.resetNonce}>{this.props.children}</React.Fragment>;
  }
}