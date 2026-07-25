import './domPolyfills';
import * as React from 'react';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';

export interface IRenderResult {
  container: HTMLElement;
  unmount: () => void;
  rerender: (next: React.ReactElement) => void;
  text: () => string;
  queryByText: (needle: string) => HTMLElement | null;
  getAllByRole: (role: string) => HTMLElement[];
  click: (element: HTMLElement) => void;
}

export function renderComponent(element: React.ReactElement): IRenderResult {
  const container: HTMLElement = document.createElement('div');
  document.body.appendChild(container);
  let root: Root | undefined;

  act(() => {
    root = createRoot(container);
    root.render(element);
  });

  return {
    container,
    unmount: () => {
      act(() => {
        root?.unmount();
      });
      container.remove();
    },
    rerender: (next: React.ReactElement) => {
      act(() => {
        root?.render(next);
      });
    },
    text: () => container.textContent ?? '',
    queryByText: (needle: string) => {
      const all: HTMLElement[] = Array.from(container.querySelectorAll<HTMLElement>('*'));
      return (
        all.filter((el) => {
          const own: string = Array.from(el.childNodes)
            .filter((n) => n.nodeType === Node.TEXT_NODE)
            .map((n) => n.textContent ?? '')
            .join('');
          return own.indexOf(needle) !== -1;
        })[0] ?? null
      );
    },
    getAllByRole: (role: string) =>
      Array.from(container.querySelectorAll<HTMLElement>(`[role="${role}"]`)),
    click: (element: HTMLElement) => {
      act(() => {
        element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
      });
    }
  };
}
