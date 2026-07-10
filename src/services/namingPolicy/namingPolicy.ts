import type {
  AvailabilityChecker,
  INamingResult,
  IRejectedCandidate,
  UpnAvailability
} from '../../models';

/**
 * Naming policy engine (spec Section 5). Everything in this file is pure —
 * directory lookups are injected via AvailabilityChecker so the algorithm
 * is unit-testable without Graph.
 */

const CHAR_MAP: Record<string, string> = {
  'å': 'a', // å
  'ä': 'a', // ä
  'ö': 'o', // ö
  'é': 'e', // é
  'ü': 'u', // ü
  'ø': 'o', // ø
  'æ': 'ae', // æ
  'ß': 'ss', // ß
  'þ': 'th', // þ
  'ð': 'd', // ð
  'đ': 'd', // đ
  'ł': 'l' // ł
};

const COMBINING_MARK_LOW: number = 0x300;
const COMBINING_MARK_HIGH: number = 0x36f;

function stripCombiningMarks(decomposed: string): string {
  let result: string = '';
  for (const ch of decomposed) {
    const code: number = ch.charCodeAt(0);
    if (code < COMBINING_MARK_LOW || code > COMBINING_MARK_HIGH) {
      result += ch;
    }
  }
  return result;
}

/** Lowercase, transliterate diacritics, strip to [a-z0-9.-], collapse repeated separators. */
export function normalizeNamePart(input: string): string {
  const lowered: string = (input ?? '').toLowerCase();
  let value: string = '';
  for (const ch of lowered) {
    if (CHAR_MAP[ch] !== undefined) {
      value += CHAR_MAP[ch];
    } else if (ch.charCodeAt(0) > 0x7f) {
      // Generic fallback: decompose and drop combining marks (e.g. n with tilde).
      value += stripCombiningMarks(ch.normalize('NFD'));
    } else {
      value += ch;
    }
  }
  value = value.replace(/[^a-z0-9.-]/g, '');
  value = value.replace(/\.{2,}/g, '.').replace(/-{2,}/g, '-');
  // Strip leading/trailing separators (e.g. a last name typed "-Smith")
  // here, at the source, so every downstream candidate already satisfies
  // the UPN_LOCAL shape instead of failing validation only much later.
  value = value.replace(/^[.-]+/, '').replace(/[.-]+$/, '');
  return value;
}

/** mailNickname rules: max 64 chars, no leading/trailing dot or hyphen. */
export function toMailNickname(localPart: string): string {
  return (localPart ?? '').slice(0, 64).replace(/^[.-]+/, '').replace(/[.-]+$/, '');
}

/**
 * Candidate order: first.last → f.last → first.l → first.lastN.
 * Returns the three fixed candidates; suffixed candidates are produced
 * on demand via `suffixedCandidate`.
 */
export function buildBaseCandidates(firstName: string, lastName: string): string[] {
  const first: string = normalizeNamePart(firstName);
  const last: string = normalizeNamePart(lastName);
  if (!first || !last) {
    const single: string = toMailNickname(first || last);
    return single ? [single] : [];
  }
  const candidates: string[] = [
    `${first}.${last}`,
    `${first.charAt(0)}.${last}`,
    `${first}.${last.charAt(0)}`
  ];
  // Preserve order while removing duplicates (single-letter names collapse candidates).
  const unique: string[] = [];
  for (const c of candidates) {
    const nick: string = toMailNickname(c);
    if (nick && unique.indexOf(nick) === -1) {
      unique.push(nick);
    }
  }
  return unique;
}

export function suffixedCandidate(firstName: string, lastName: string, n: number): string {
  const first: string = normalizeNamePart(firstName);
  const last: string = normalizeNamePart(lastName);
  const suffix: string = String(n);
  const base: string = first && last ? `${first}.${last}` : first || last;
  // Keep the suffix intact when truncating to the 64-char budget.
  const room: number = 64 - suffix.length;
  return toMailNickname(base.slice(0, room) + suffix);
}

export function isReserved(localPart: string, reservedNames: readonly string[]): boolean {
  return reservedNames.indexOf(localPart) !== -1;
}

export interface IResolveNamingInput {
  firstName: string;
  lastName: string;
  domain: string;
  reservedNames: readonly string[];
  checkAvailability: AvailabilityChecker;
  /** How many suffixed candidates to try before giving up. */
  maxSuffix?: number;
  /** How many spare available candidates to surface. */
  alternativesWanted?: number;
  signal?: AbortSignal;
}

/**
 * Resolve the UPN local part: first available candidate wins; rejected
 * candidates keep their reason; up to 3 further available candidates are
 * offered as alternatives the operator may pick instead.
 */
export async function resolveNaming(input: IResolveNamingInput): Promise<INamingResult> {
  const maxSuffix: number = input.maxSuffix ?? 50;
  const alternativesWanted: number = input.alternativesWanted ?? 3;
  const rejected: IRejectedCandidate[] = [];
  const alternatives: string[] = [];
  let chosen: string | null = null;

  const tried: Set<string> = new Set();

  const evaluate = async (candidate: string): Promise<boolean> => {
    if (!candidate || tried.has(candidate)) {
      return false;
    }
    tried.add(candidate);
    if (isReserved(candidate, input.reservedNames)) {
      rejected.push({ candidate, reason: 'reserved' });
      return false;
    }
    const availability: UpnAvailability = await input.checkAvailability(
      `${candidate}@${input.domain}`,
      input.signal
    );
    if (availability === 'taken') {
      rejected.push({ candidate, reason: 'collision' });
      return false;
    }
    if (availability === 'taken-soft-deleted') {
      rejected.push({ candidate, reason: 'collision-soft-deleted' });
      return false;
    }
    if (chosen === null) {
      chosen = candidate;
    } else if (alternatives.length < alternativesWanted) {
      alternatives.push(candidate);
    }
    return chosen !== null && alternatives.length >= alternativesWanted;
  };

  for (const candidate of buildBaseCandidates(input.firstName, input.lastName)) {
    if (await evaluate(candidate)) {
      return { chosen, rejected, alternatives };
    }
  }
  for (let n = 2; n <= maxSuffix; n++) {
    if (await evaluate(suffixedCandidate(input.firstName, input.lastName, n))) {
      return { chosen, rejected, alternatives };
    }
  }
  return { chosen, rejected, alternatives };
}
