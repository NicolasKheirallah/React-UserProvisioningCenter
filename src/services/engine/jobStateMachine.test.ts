import { assertTransition, canStartJob, canTransition, isTerminal } from './jobStateMachine';
import type { JobStatus } from '../../models';

const ALL: JobStatus[] = [
  'Draft',
  'PendingApproval',
  'Approved',
  'Scheduled',
  'Running',
  'PartiallyFailed',
  'Failed',
  'Completed',
  'Cancelled'
];

describe('jobStateMachine', () => {
  it('follows the happy path Draft → ... → Completed', () => {
    expect(canTransition('Draft', 'PendingApproval')).toBe(true);
    expect(canTransition('PendingApproval', 'Approved')).toBe(true);
    expect(canTransition('Approved', 'Running')).toBe(true);
    expect(canTransition('Approved', 'Scheduled')).toBe(true);
    expect(canTransition('Scheduled', 'Running')).toBe(true);
    expect(canTransition('Running', 'Completed')).toBe(true);
  });

  it('refuses to run unapproved jobs (Phase 1 acceptance)', () => {
    expect(canStartJob('Draft')).toBe(false);
    expect(canStartJob('PendingApproval')).toBe(false);
    expect(canStartJob('Cancelled')).toBe(false);
    expect(canStartJob('Completed')).toBe(false);
    expect(canStartJob('Approved')).toBe(true);
    expect(canStartJob('Scheduled')).toBe(true);
    expect(canStartJob('Running')).toBe(true);
    expect(canStartJob('PartiallyFailed')).toBe(true);
    expect(canStartJob('Failed')).toBe(true);
  });

  it('never skips approval', () => {
    expect(canTransition('Draft', 'Approved')).toBe(false);
    expect(canTransition('Draft', 'Running')).toBe(false);
    expect(canTransition('PendingApproval', 'Running')).toBe(false);
    expect(canTransition('PendingApproval', 'Scheduled')).toBe(false);
  });

  it('supports manual retry from PartiallyFailed and Failed', () => {
    expect(canTransition('PartiallyFailed', 'Running')).toBe(true);
    expect(canTransition('Failed', 'Running')).toBe(true);
  });

  it('treats Completed and Cancelled as terminal', () => {
    for (const target of ALL) {
      expect(canTransition('Completed', target)).toBe(false);
      expect(canTransition('Cancelled', target)).toBe(false);
    }
    expect(isTerminal('Completed')).toBe(true);
    expect(isTerminal('Cancelled')).toBe(true);
    expect(isTerminal('Running')).toBe(false);
  });

  it('allows cancellation from every non-terminal state', () => {
    for (const from of ALL) {
      if (!isTerminal(from)) {
        expect(canTransition(from, 'Cancelled')).toBe(true);
      }
    }
  });

  it('assertTransition throws on illegal moves', () => {
    expect(() => assertTransition('Draft', 'Running')).toThrow(/Invalid job transition/);
    expect(() => assertTransition('Approved', 'Running')).not.toThrow();
  });
});
