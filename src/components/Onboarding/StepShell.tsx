import * as React from 'react';
import { Caption1, Subtitle2, makeStyles, tokens } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalL,
    maxWidth: '600px'
  },
  wide: {
    maxWidth: '760px'
  },
  header: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalXS
  },
  description: {
    color: tokens.colorNeutralForeground3
  }
});

export interface IStepShellProps {
  title: string;
  description?: string;
  /** Tables and grids get a wider column than plain forms. */
  wide?: boolean;
  children?: React.ReactNode;
}

/**
 * Standard wizard step column: Subtitle2 step title, Caption1 description,
 * content constrained to a readable single column (Fluent 2 layout guidance —
 * admin center forms never span the full page width).
 */
export const StepShell: React.FC<IStepShellProps> = (props) => {
  const styles = useStyles();
  return (
    <div className={props.wide ? `${styles.root} ${styles.wide}` : styles.root}>
      <div className={styles.header}>
        <Subtitle2 as="h4" block>
          {props.title}
        </Subtitle2>
        {props.description ? (
          <Caption1 block className={styles.description}>
            {props.description}
          </Caption1>
        ) : undefined}
      </div>
      {props.children}
    </div>
  );
};
