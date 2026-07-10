/**
 * Client-side temporary password generation. The value lives in memory only,
 * is shown once in a copy-once dialog and is never written to any list, log
 * or state (spec Sections 1 and 6).
 */

const LOWER: string = 'abcdefghjkmnpqrstuvwxyz';
const UPPER: string = 'ABCDEFGHJKMNPQRSTUVWXYZ';
const DIGITS: string = '23456789';
const SYMBOLS: string = '!@#$%^&*-_=+';
const ALL: string = LOWER + UPPER + DIGITS + SYMBOLS;

/** Uniform random integer in [0, maxExclusive) using rejection sampling (no modulo bias). */
function randomInt(maxExclusive: number): number {
  const maxUint32: number = 0x100000000;
  const limit: number = maxUint32 - (maxUint32 % maxExclusive);
  const buffer: Uint32Array = new Uint32Array(1);
  for (;;) {
    crypto.getRandomValues(buffer);
    if (buffer[0] < limit) {
      return buffer[0] % maxExclusive;
    }
  }
}

function pick(charset: string): string {
  return charset.charAt(randomInt(charset.length));
}

/** 16 chars, guaranteed upper/lower/digit/symbol, ambiguous glyphs excluded. */
export function generateTempPassword(length: number = 16): string {
  const required: string[] = [pick(LOWER), pick(UPPER), pick(DIGITS), pick(SYMBOLS)];
  const remaining: number = Math.max(length, 8) - required.length;
  const chars: string[] = [...required];
  for (let i = 0; i < remaining; i++) {
    chars.push(ALL.charAt(randomInt(ALL.length)));
  }
  // Fisher-Yates shuffle with crypto randomness so required classes aren't positional.
  for (let i = chars.length - 1; i > 0; i--) {
    const j: number = randomInt(i + 1);
    const tmp: string = chars[i];
    chars[i] = chars[j];
    chars[j] = tmp;
  }
  return chars.join('');
}