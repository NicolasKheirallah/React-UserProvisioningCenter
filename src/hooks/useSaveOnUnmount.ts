import * as React from 'react';

/**
 * Runs `save` when the component unmounts. Wizard steps unmount on every
 * navigation (Back as well as Next), so this persists in-progress form
 * values into the wizard reducer without requiring validation to pass —
 * validation re-runs when the operator moves forward again.
 */
export function useSaveOnUnmount(save: () => void): void {
  const saveRef = React.useRef(save);
  saveRef.current = save;
  React.useEffect(() => {
    return () => saveRef.current();
  }, []);
}
