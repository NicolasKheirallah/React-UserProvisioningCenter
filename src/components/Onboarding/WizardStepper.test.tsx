jest.mock(
  'UpcStrings',
  () => ({
    WizardStepsAriaLabel: 'Wizard steps',
    WizardStepOfLabel: 'Step {0} of {1}',
    WizardStepHasErrorsAria: 'has errors',
    WizardStepCompletedAria: 'completed'
  }),
  { virtual: true }
);

import * as React from 'react';
import { renderComponent, type IRenderResult } from '../../testing/renderComponent';
import { WizardStepper } from './WizardStepper';

const LABELS: string[] = ['Personal', 'Employment', 'Identity', 'Account'];

describe('WizardStepper', () => {
  let rendered: IRenderResult | undefined;

  afterEach(() => {
    rendered?.unmount();
    rendered = undefined;
  });

  it('renders every step label', () => {
    rendered = renderComponent(<WizardStepper labels={LABELS} current={0} onStepClick={() => undefined} />);
    for (const label of LABELS) {
      expect(rendered.text()).toContain(label);
    }
  });

  it('reports the step the operator is on to assistive technology', () => {
    rendered = renderComponent(<WizardStepper labels={LABELS} current={2} onStepClick={() => undefined} />);
    const current = rendered.container.querySelector('[aria-current="step"]');
    expect(current).not.toBeNull();
    expect(current?.textContent).toContain('Identity');
  });

  it('invokes onStepClick with the index that was activated', () => {
    const clicks: number[] = [];
    rendered = renderComponent(
      <WizardStepper labels={LABELS} current={3} onStepClick={(i) => clicks.push(i)} />
    );
    const buttons = rendered.container.querySelectorAll<HTMLElement>('button');
    rendered.click(buttons[1]);
    expect(clicks).toEqual([1]);
  });

  it('marks a completed-but-invalid step so the operator can find the problem', () => {
    rendered = renderComponent(
      <WizardStepper
        labels={LABELS}
        current={3}
        onStepClick={() => undefined}
        errorSteps={new Set<number>([1])}
      />
    );
    const labelled = Array.from(
      rendered.container.querySelectorAll<HTMLElement>('[aria-label]')
    ).map((el) => el.getAttribute('aria-label') ?? '');
    expect(labelled.some((l) => l.indexOf('has errors') !== -1)).toBe(true);
    expect(labelled.some((l) => l.indexOf('Employment (has errors)') !== -1)).toBe(true);
  });

  it('does not announce errors when every completed step is valid', () => {
    rendered = renderComponent(
      <WizardStepper labels={LABELS} current={3} onStepClick={() => undefined} errorSteps={new Set<number>()} />
    );
    const labelled = Array.from(
      rendered.container.querySelectorAll<HTMLElement>('[aria-label]')
    ).map((el) => el.getAttribute('aria-label') ?? '');
    expect(labelled.some((l) => l.indexOf('has errors') !== -1)).toBe(false);
    expect(labelled.some((l) => l.indexOf('Employment (completed)') !== -1)).toBe(true);
  });
});
