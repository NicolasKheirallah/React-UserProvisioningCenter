import * as React from 'react';
import {
  Toast,
  ToastTitle,
  Toaster,
  useId,
  useToastController,
  type ToastIntent
} from '@fluentui/react-components';

const ToasterIdContext: React.Context<string> = React.createContext<string>('upc-toaster');

/**
 * One Toaster per web part instance (id via useId so multiple instances on a
 * page stay independent), shared through context so any component can raise
 * transient confirmations.
 */
export const AppToasterProvider: React.FC<{ children?: React.ReactNode }> = ({ children }) => {
  const toasterId: string = useId('upc-toaster');
  return (
    <ToasterIdContext.Provider value={toasterId}>
      {children}
      <Toaster toasterId={toasterId} position="top-end" />
    </ToasterIdContext.Provider>
  );
};

/** Dispatches a single-line toast; the default intent is success. */
export function useAppToast(): (message: string, intent?: ToastIntent) => void {
  const toasterId: string = React.useContext(ToasterIdContext);
  const { dispatchToast } = useToastController(toasterId);
  return React.useCallback(
    (message: string, intent: ToastIntent = 'success') => {
      dispatchToast(
        <Toast>
          <ToastTitle>{message}</ToastTitle>
        </Toast>,
        { intent }
      );
    },
    [dispatchToast]
  );
}
