import { buildJobFilter } from './jobFilter';

describe('buildJobFilter', () => {
  it('returns an empty filter for an undefined or empty query', () => {
    expect(buildJobFilter(undefined)).toBe('');
    expect(buildJobFilter({})).toBe('');
  });

  it('searches target upn and job id together', () => {
    expect(buildJobFilter({ search: 'anna' })).toBe(
      "(startswith(TargetUpn,'anna') or startswith(Title,'anna'))"
    );
  });

  it('ORs multiple statuses and job types, ANDing the groups', () => {
    const filter = buildJobFilter({ status: ['Failed', 'Running'], jobType: ['Onboard'] });
    expect(filter).toBe(
      "(Status eq 'Failed' or Status eq 'Running') and (JobType eq 'Onboard')"
    );
  });

  it('filters by exact batch id so a bulk submission can be rolled up', () => {
    expect(buildJobFilter({ batchId: 'batch-guid-1' })).toBe("BatchId eq 'batch-guid-1'");
  });

  it('escapes single quotes to prevent OData injection', () => {
    expect(buildJobFilter({ search: "o'brien" })).toContain("startswith(TargetUpn,'o''brien')");
    expect(buildJobFilter({ batchId: "x'y" })).toBe("BatchId eq 'x''y'");
  });

  it('expands date-only bounds to a full-day range', () => {
    const filter = buildJobFilter({ createdFromUtc: '2026-07-01', createdToUtc: '2026-07-31' });
    expect(filter).toContain("Created ge datetime'2026-07-01T00:00:00.000Z'");
    expect(filter).toContain("Created le datetime'2026-07-31T23:59:59.999Z'");
  });

  it('ignores unparseable dates rather than emitting a broken filter', () => {
    expect(buildJobFilter({ createdFromUtc: 'nonsense' })).toBe('');
  });

  it('ignores blank batch id and search', () => {
    expect(buildJobFilter({ batchId: '   ', search: '  ' })).toBe('');
  });
});
