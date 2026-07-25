import { isJobStalled, JOB_STALE_AFTER_MS } from './IJobQuery';

describe('isJobStalled', () => {
  const now: number = new Date('2026-07-25T12:00:00Z').getTime();

  it('flags a Running job whose lock is older than the stale threshold', () => {
    const runningSince: string = new Date(now - JOB_STALE_AFTER_MS - 1000).toISOString();
    expect(isJobStalled({ status: 'Running', runningSince }, now)).toBe(true);
  });

  it('does not flag a Running job that started recently', () => {
    const runningSince: string = new Date(now - 60_000).toISOString();
    expect(isJobStalled({ status: 'Running', runningSince }, now)).toBe(false);
  });

  it('does not flag a job that is not Running even with an old lock', () => {
    const runningSince: string = new Date(now - JOB_STALE_AFTER_MS - 1000).toISOString();
    expect(isJobStalled({ status: 'Completed', runningSince }, now)).toBe(false);
    expect(isJobStalled({ status: 'Failed', runningSince }, now)).toBe(false);
  });

  it('does not flag a Running job with no lock timestamp', () => {
    expect(isJobStalled({ status: 'Running', runningSince: null }, now)).toBe(false);
  });

  it('does not flag when the timestamp is unparseable', () => {
    expect(isJobStalled({ status: 'Running', runningSince: 'not-a-date' }, now)).toBe(false);
  });

  it('honours a custom stale threshold', () => {
    const runningSince: string = new Date(now - 5000).toISOString();
    expect(isJobStalled({ status: 'Running', runningSince }, now, 1000)).toBe(true);
    expect(isJobStalled({ status: 'Running', runningSince }, now, 10_000)).toBe(false);
  });
});
