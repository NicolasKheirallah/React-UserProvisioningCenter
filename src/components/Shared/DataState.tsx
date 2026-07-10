import * as React from 'react';
import {
  Button,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  Skeleton,
  SkeletonItem,
  Text,
  makeStyles,
  tokens
} from '@fluentui/react-components';
import * as strings from 'UpcStrings';
import { GraphServiceError } from '../../services/graph/GraphError';

const useStyles = makeStyles({
  skeleton: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM
  },
  empty: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    rowGap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalXXL,
    color: tokens.colorNeutralForeground3
  }
});

export interface IDataStateProps {
  isLoading: boolean;
  error?: unknown;
  isEmpty?: boolean;
  emptyTitle?: string;
  emptyBody?: string;
  /** Optional call to action rendered under the empty-state text. */
  emptyAction?: React.ReactNode;
  onRetry?: () => void;
  children?: React.ReactNode;
}

/**
 * Standard four-state wrapper (spec Section 3): Loading (Skeleton),
 * Error (actionable + correlation id), Empty (instructional), Data.
 */
export const DataState: React.FC<IDataStateProps> = (props) => {
  const styles = useStyles();

  if (props.isLoading) {
    return (
      <Skeleton aria-label={strings.LoadingLabel} className={styles.skeleton}>
        <SkeletonItem size={24} />
        <SkeletonItem size={24} />
        <SkeletonItem size={24} />
      </Skeleton>
    );
  }

  if (props.error) {
    const correlationId: string =
      props.error instanceof GraphServiceError ? props.error.requestId : '';
    return (
      <MessageBar intent="error" layout="multiline">
        <MessageBarBody>
          <MessageBarTitle>{strings.ErrorGenericTitle}</MessageBarTitle>
          {props.error instanceof Error ? props.error.message : String(props.error)}
          {correlationId ? (
            <>
              {' '}
              <Text size={200}>
                {strings.CorrelationIdLabel}: {correlationId}
              </Text>
            </>
          ) : undefined}
        </MessageBarBody>
        {props.onRetry ? (
          <MessageBarActions>
            <Button onClick={props.onRetry}>{strings.RetryLabel}</Button>
          </MessageBarActions>
        ) : undefined}
      </MessageBar>
    );
  }

  if (props.isEmpty) {
    return (
      <div className={styles.empty}>
        <Text weight="semibold">{props.emptyTitle ?? strings.EmptyGenericTitle}</Text>
        <Text size={200}>{props.emptyBody ?? strings.EmptyGenericBody}</Text>
        {props.emptyAction}
      </div>
    );
  }

  return <>{props.children}</>;
};
