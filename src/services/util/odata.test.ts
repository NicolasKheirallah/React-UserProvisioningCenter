import { escapeODataLiteral } from './odata';

describe('escapeODataLiteral', () => {
  it('returns plain strings unchanged', () => {
    expect(escapeODataLiteral('hello')).toBe('hello');
  });

  it('doubles single quotes', () => {
    expect(escapeODataLiteral("O'Brien")).toBe("O''Brien");
  });

  it('handles multiple quotes', () => {
    expect(escapeODataLiteral("a'b'c")).toBe("a''b''c");
  });

  it('handles empty string', () => {
    expect(escapeODataLiteral('')).toBe('');
  });

  it('prevents OData injection via quote escaping', () => {
    const malicious = "admin' or 1 eq 1";
    const escaped = escapeODataLiteral(malicious);
    expect(escaped).toBe("admin'' or 1 eq 1");
  });
});