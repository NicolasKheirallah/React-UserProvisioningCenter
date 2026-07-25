import { buildAuditFilter } from './auditFilter';

describe('buildAuditFilter', () => {
  it('returns an empty filter for an undefined or empty query', () => {
    expect(buildAuditFilter(undefined)).toBe('');
    expect(buildAuditFilter({})).toBe('');
  });

  it('filters by exact job id', () => {
    expect(buildAuditFilter({ jobId: 'job-1' })).toBe("JobId eq 'job-1'");
  });

  it('uses startswith for actor and target so partial names match', () => {
    expect(buildAuditFilter({ actor: 'anna' })).toBe("startswith(Actor,'anna')");
    expect(buildAuditFilter({ targetUser: 'bob@x.com' })).toBe("startswith(TargetUser,'bob@x.com')");
  });

  it('ORs multiple results and ANDs distinct clauses', () => {
    const filter = buildAuditFilter({ result: ['Failure', 'Skipped'], jobId: 'job-9' });
    expect(filter).toBe("JobId eq 'job-9' and (Result eq 'Failure' or Result eq 'Skipped')");
  });

  it('escapes single quotes to prevent OData injection', () => {
    const filter = buildAuditFilter({ actor: "o'brien" });
    expect(filter).toBe("startswith(Actor,'o''brien')");
  });

  it('expands date-only bounds to full-day range', () => {
    const filter = buildAuditFilter({ fromUtc: '2026-07-01', toUtc: '2026-07-31' });
    expect(filter).toContain("TimestampUtc ge datetime'2026-07-01T00:00:00.000Z'");
    expect(filter).toContain("TimestampUtc le datetime'2026-07-31T23:59:59.999Z'");
  });

  it('ignores unparseable dates rather than emitting a broken filter', () => {
    expect(buildAuditFilter({ fromUtc: 'not-a-date' })).toBe('');
  });

  it('ignores blank strings', () => {
    expect(buildAuditFilter({ jobId: '   ', actor: '', targetUser: '  ' })).toBe('');
  });
});
