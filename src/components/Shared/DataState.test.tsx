jest.mock(
  'UpcStrings',
  () => ({
    LoadingLabel: 'Loading',
    ErrorGenericTitle: 'Something went wrong',
    CorrelationIdLabel: 'Correlation id',
    RetryLabel: 'Try again',
    EmptyGenericTitle: 'Nothing here yet',
    EmptyGenericBody: 'No items to show.',
    SlowLoadTitle: 'This is taking longer than usual',
    SlowLoadBody: 'The request has not completed yet.'
  }),
  { virtual: true }
);

import * as React from 'react';
import { renderComponent, type IRenderResult } from '../../testing/renderComponent';
import { DataState } from './DataState';


describe('DataState', () => {
  let rendered: IRenderResult | undefined;

  afterEach(() => {
    rendered?.unmount();
    rendered = undefined;
    jest.useRealTimers();
  });

  it('renders children when there is data', () => {
    rendered = renderComponent(
      <DataState isLoading={false}>
        <p>the payload</p>
      </DataState>
    );
    expect(rendered.text()).toContain('the payload');
  });

  it('shows an error instead of children when the query failed', () => {
    rendered = renderComponent(
      <DataState isLoading={false} error={new Error('Column TargetUpn does not exist')}>
        <p>the payload</p>
      </DataState>
    );
    expect(rendered.text()).toContain('Column TargetUpn does not exist');
    expect(rendered.text()).not.toContain('the payload');
  });

  it('prefers the error over the loading skeleton so a failure is never masked', () => {
    rendered = renderComponent(
      <DataState isLoading error={new Error('boom')}>
        <p>the payload</p>
      </DataState>
    );
    expect(rendered.text()).toContain('boom');
  });

  it('shows the empty state only when there is no error and nothing to render', () => {
    rendered = renderComponent(<DataState isLoading={false} isEmpty />);
    expect(rendered.text()).toContain('Nothing here yet');
  });

  it('does not warn about a slow load before the threshold elapses', () => {
    jest.useFakeTimers();
    rendered = renderComponent(<DataState isLoading slowAfterMs={15_000} />);
    expect(rendered.text()).not.toContain('This is taking longer than usual');
  });

  it('warns and offers a retry once a load exceeds the slow threshold', () => {
    jest.useFakeTimers();
    let retried: number = 0;
    rendered = renderComponent(
      <DataState isLoading slowAfterMs={15_000} onRetry={() => (retried += 1)} />
    );

    React.act(() => {
      jest.advanceTimersByTime(15_001);
    });

    expect(rendered.text()).toContain('This is taking longer than usual');
    const retry = rendered.queryByText('Try again');
    expect(retry).not.toBeNull();
    rendered.click(retry as HTMLElement);
    expect(retried).toBe(1);
  });

  it('clears the slow warning once loading finishes', () => {
    jest.useFakeTimers();
    rendered = renderComponent(<DataState isLoading slowAfterMs={15_000} />);
    React.act(() => {
      jest.advanceTimersByTime(15_001);
    });
    expect(rendered.text()).toContain('This is taking longer than usual');

    rendered.rerender(
      <DataState isLoading={false} slowAfterMs={15_000}>
        <p>arrived</p>
      </DataState>
    );
    expect(rendered.text()).toContain('arrived');
    expect(rendered.text()).not.toContain('This is taking longer than usual');
  });
});
