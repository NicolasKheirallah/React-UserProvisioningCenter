import {
  buildBaseCandidates,
  isReserved,
  normalizeNamePart,
  resolveNaming,
  suffixedCandidate,
  toMailNickname
} from './namingPolicy';
import type { AvailabilityChecker, UpnAvailability } from '../../models';
import { DEFAULT_RESERVED_NAMES } from '../../constants/reservedNames';

function checkerFor(map: Record<string, UpnAvailability>): AvailabilityChecker {
  return async (upn: string): Promise<UpnAvailability> => map[upn] ?? 'available';
}

describe('normalizeNamePart', () => {
  it('lowercases and transliterates Nordic diacritics', () => {
    expect(normalizeNamePart('Åsa')).toBe('asa');
    expect(normalizeNamePart('Öström')).toBe('ostrom');
    expect(normalizeNamePart('Sörensen')).toBe('sorensen');
    expect(normalizeNamePart('Ægir')).toBe('aegir');
    expect(normalizeNamePart('Ørsted')).toBe('orsted');
    expect(normalizeNamePart('Müller')).toBe('muller');
    expect(normalizeNamePart('René')).toBe('rene');
  });

  it('handles diacritics outside the explicit map via decomposition', () => {
    expect(normalizeNamePart('Peña')).toBe('pena');
    expect(normalizeNamePart('François')).toBe('francois');
  });

  it('strips characters outside [a-z0-9.-]', () => {
    expect(normalizeNamePart("O'Brien")).toBe('obrien');
    expect(normalizeNamePart('van der Berg')).toBe('vanderberg');
    expect(normalizeNamePart('smith_jones!')).toBe('smithjones');
  });

  it('collapses repeated separators', () => {
    expect(normalizeNamePart('a..b')).toBe('a.b');
    expect(normalizeNamePart('a--b')).toBe('a-b');
  });

  it('strips a leading or trailing hyphen introduced by a mistyped name', () => {

    expect(normalizeNamePart('-Smith')).toBe('smith');
    expect(normalizeNamePart('Smith-')).toBe('smith');
    expect(normalizeNamePart('-Smith-')).toBe('smith');
    expect(normalizeNamePart('.Smith.')).toBe('smith');
  });
});

describe('toMailNickname', () => {
  it('caps at 64 characters', () => {
    expect(toMailNickname('x'.repeat(70))).toHaveLength(64);
  });

  it('strips leading and trailing dots', () => {
    expect(toMailNickname('.anna.svensson.')).toBe('anna.svensson');
  });

  it('strips leading and trailing hyphens', () => {
    expect(toMailNickname('-anna.svensson-')).toBe('anna.svensson');
  });
});

describe('buildBaseCandidates', () => {
  it('never produces a candidate with a leading/trailing hyphen or dot (must satisfy UPN_LOCAL)', () => {

    const UPN_LOCAL: RegExp = /^[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?$/;
    for (const candidate of buildBaseCandidates('-Anna', 'Svensson-')) {
      expect(candidate).toMatch(UPN_LOCAL);
    }
  });

  it('produces candidates in the mandated order', () => {
    expect(buildBaseCandidates('Anna', 'Svensson')).toEqual([
      'anna.svensson',
      'a.svensson',
      'anna.s'
    ]);
  });

  it('deduplicates when names are single letters', () => {
    expect(buildBaseCandidates('A', 'B')).toEqual(['a.b']);
  });

  it('copes with a missing name part', () => {
    expect(buildBaseCandidates('', 'Svensson')).toEqual(['svensson']);
  });
});

describe('suffixedCandidate', () => {
  it('appends the increment', () => {
    expect(suffixedCandidate('Anna', 'Svensson', 2)).toBe('anna.svensson2');
  });

  it('keeps the suffix when truncating to 64 chars', () => {
    const result: string = suffixedCandidate('a'.repeat(40), 'b'.repeat(40), 12);
    expect(result).toHaveLength(64);
    expect(result.endsWith('12')).toBe(true);
  });
});

describe('isReserved', () => {
  it('blocks the baseline reserved names', () => {
    for (const name of DEFAULT_RESERVED_NAMES) {
      expect(isReserved(name, DEFAULT_RESERVED_NAMES)).toBe(true);
    }
    expect(isReserved('anna.svensson', DEFAULT_RESERVED_NAMES)).toBe(false);
  });
});

describe('resolveNaming', () => {
  const domain: string = 'contoso.com';

  it('chooses first.last when available and offers alternatives', async () => {
    const result = await resolveNaming({
      firstName: 'Anna',
      lastName: 'Svensson',
      domain,
      reservedNames: DEFAULT_RESERVED_NAMES,
      checkAvailability: checkerFor({})
    });
    expect(result.chosen).toBe('anna.svensson');
    expect(result.alternatives).toEqual(['a.svensson', 'anna.s', 'anna.svensson2']);
    expect(result.rejected).toEqual([]);
  });

  it('falls back to f.last on collision and records the rejection', async () => {
    const result = await resolveNaming({
      firstName: 'Anna',
      lastName: 'Svensson',
      domain,
      reservedNames: DEFAULT_RESERVED_NAMES,
      checkAvailability: checkerFor({ 'anna.svensson@contoso.com': 'taken' })
    });
    expect(result.chosen).toBe('a.svensson');
    expect(result.rejected).toEqual([{ candidate: 'anna.svensson', reason: 'collision' }]);
  });

  it('treats soft-deleted users as collisions with a distinct reason', async () => {
    const result = await resolveNaming({
      firstName: 'Anna',
      lastName: 'Svensson',
      domain,
      reservedNames: DEFAULT_RESERVED_NAMES,
      checkAvailability: checkerFor({
        'anna.svensson@contoso.com': 'taken-soft-deleted'
      })
    });
    expect(result.chosen).toBe('a.svensson');
    expect(result.rejected).toEqual([
      { candidate: 'anna.svensson', reason: 'collision-soft-deleted' }
    ]);
  });

  it('blocks reserved candidates without a directory call', async () => {
    const calls: string[] = [];
    const checker: AvailabilityChecker = async (upn: string): Promise<UpnAvailability> => {
      calls.push(upn);
      return 'available';
    };
    const result = await resolveNaming({
      firstName: 'Ad',
      lastName: 'Min',
      domain,
      reservedNames: [...DEFAULT_RESERVED_NAMES, 'ad.min'],
      checkAvailability: checker
    });
    expect(result.rejected).toContainEqual({ candidate: 'ad.min', reason: 'reserved' });
    expect(calls).not.toContain('ad.min@contoso.com');
    expect(result.chosen).toBe('a.min');
  });

  it('increments the numeric suffix until an available candidate appears', async () => {
    const result = await resolveNaming({
      firstName: 'Anna',
      lastName: 'Svensson',
      domain,
      reservedNames: DEFAULT_RESERVED_NAMES,
      checkAvailability: checkerFor({
        'anna.svensson@contoso.com': 'taken',
        'a.svensson@contoso.com': 'taken',
        'anna.s@contoso.com': 'taken',
        'anna.svensson2@contoso.com': 'taken',
        'anna.svensson3@contoso.com': 'taken'
      })
    });
    expect(result.chosen).toBe('anna.svensson4');
    expect(result.rejected).toHaveLength(5);
  });

  it('returns null chosen when the search budget is exhausted', async () => {
    const result = await resolveNaming({
      firstName: 'Anna',
      lastName: 'Svensson',
      domain,
      reservedNames: DEFAULT_RESERVED_NAMES,
      checkAvailability: async (): Promise<UpnAvailability> => 'taken',
      maxSuffix: 4
    });
    expect(result.chosen).toBeNull();
    expect(result.alternatives).toEqual([]);
  });
});
