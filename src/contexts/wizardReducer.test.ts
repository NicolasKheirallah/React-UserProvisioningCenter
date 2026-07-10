import { initialWizardState, wizardReducer, WIZARD_STEP_COUNT } from './WizardContext';
import type { IWizardState } from './WizardContext';

describe('wizardReducer', () => {
  it('exposes seven steps', () => {
    expect(WIZARD_STEP_COUNT).toBe(7);
  });

  it('initial state starts at step 0 with an empty draft', () => {
    expect(initialWizardState.step).toBe(0);
    expect(initialWizardState.draft.personal.firstName).toBe('');
    expect(initialWizardState.draft.licenses).toEqual([]);
  });

  it('savePersonal keeps the other draft slices intact', () => {
    let state: IWizardState = wizardReducer(initialWizardState, {
      type: 'saveEmployment',
      employment: { ...initialWizardState.draft.employment, jobTitle: 'Controller' }
    });
    state = wizardReducer(state, {
      type: 'savePersonal',
      personal: { ...initialWizardState.draft.personal, firstName: 'Anna' }
    });
    expect(state.draft.personal.firstName).toBe('Anna');
    expect(state.draft.employment.jobTitle).toBe('Controller');
  });

  it('back/next navigation never touches saved draft values', () => {
    let state: IWizardState = wizardReducer(initialWizardState, {
      type: 'savePersonal',
      personal: { ...initialWizardState.draft.personal, firstName: 'Anna', lastName: 'Svensson' }
    });
    state = wizardReducer(state, { type: 'next' });
    state = wizardReducer(state, { type: 'back' });
    expect(state.step).toBe(0);
    expect(state.draft.personal.firstName).toBe('Anna');
    expect(state.draft.personal.lastName).toBe('Svensson');
  });

  it('goto jumps to a step and clamps out-of-range targets', () => {
    let state: IWizardState = wizardReducer(initialWizardState, { type: 'goto', step: 3 });
    expect(state.step).toBe(3);
    state = wizardReducer(state, { type: 'goto', step: -2 });
    expect(state.step).toBe(0);
    state = wizardReducer(state, { type: 'goto', step: 99 });
    expect(state.step).toBe(6);
  });

  it('clamps navigation at both ends', () => {
    expect(wizardReducer(initialWizardState, { type: 'back' }).step).toBe(0);
    let state: IWizardState = initialWizardState;
    for (let i = 0; i < 10; i++) {
      state = wizardReducer(state, { type: 'next' });
    }
    expect(state.step).toBe(6);
  });

  it('toggleLicense adds and removes a selection', () => {
    const license = { skuId: 'sku-1', skuPartNumber: 'SPE_E5' };
    let state: IWizardState = wizardReducer(initialWizardState, { type: 'toggleLicense', license });
    expect(state.draft.licenses).toHaveLength(1);
    state = wizardReducer(state, { type: 'toggleLicense', license });
    expect(state.draft.licenses).toHaveLength(0);
  });

  it('reset restores the initial draft and step', () => {
    let state: IWizardState = wizardReducer(initialWizardState, {
      type: 'savePersonal',
      personal: { ...initialWizardState.draft.personal, firstName: 'Anna' }
    });
    state = wizardReducer(state, { type: 'next' });
    state = wizardReducer(state, { type: 'reset' });
    expect(state.step).toBe(0);
    expect(state.draft.personal.firstName).toBe('');
  });
});
