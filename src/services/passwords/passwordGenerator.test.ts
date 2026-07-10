import { generateTempPassword } from './passwordGenerator';

describe('generateTempPassword', () => {
  it('produces the requested length', () => {
    expect(generateTempPassword(16)).toHaveLength(16);
    expect(generateTempPassword(20)).toHaveLength(20);
  });

  it('never goes below 8 characters', () => {
    expect(generateTempPassword(4).length).toBeGreaterThanOrEqual(8);
  });

  it('contains all four character classes', () => {
    for (let i = 0; i < 25; i++) {
      const pw: string = generateTempPassword();
      expect(pw).toMatch(/[a-z]/);
      expect(pw).toMatch(/[A-Z]/);
      expect(pw).toMatch(/[0-9]/);
      expect(pw).toMatch(/[!@#$%^&*\-_=+]/);
    }
  });

  it('excludes ambiguous glyphs', () => {
    for (let i = 0; i < 25; i++) {
      expect(generateTempPassword()).not.toMatch(/[O0Il1]/);
    }
  });

  it('does not repeat across invocations', () => {
    const seen: Set<string> = new Set();
    for (let i = 0; i < 50; i++) {
      seen.add(generateTempPassword());
    }
    expect(seen.size).toBe(50);
  });
});
