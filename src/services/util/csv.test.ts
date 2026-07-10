import { parseCsv, toCsv } from './csv';

describe('parseCsv', () => {
  it('parses a simple comma-separated grid', () => {
    expect(parseCsv('a,b,c\n1,2,3')).toEqual([
      ['a', 'b', 'c'],
      ['1', '2', '3']
    ]);
  });

  it('handles quoted fields containing commas', () => {
    expect(parseCsv('name,note\n"Svensson, Anna",ok')).toEqual([
      ['name', 'note'],
      ['Svensson, Anna', 'ok']
    ]);
  });

  it('unescapes doubled quotes inside a quoted field', () => {
    expect(parseCsv('note\n"She said ""hi"""')).toEqual([['note'], ['She said "hi"']]);
  });

  it('handles CRLF and LF line endings the same way', () => {
    expect(parseCsv('a,b\r\n1,2\r\n3,4')).toEqual([
      ['a', 'b'],
      ['1', '2'],
      ['3', '4']
    ]);
  });

  it('drops blank lines', () => {
    expect(parseCsv('a,b\n\n1,2\n\n')).toEqual([
      ['a', 'b'],
      ['1', '2']
    ]);
  });

  it('preserves an empty cell between two commas', () => {
    expect(parseCsv('a,,c')).toEqual([['a', '', 'c']]);
  });

  it('returns an empty array for empty input', () => {
    expect(parseCsv('')).toEqual([]);
  });

  it('handles a file with no trailing newline', () => {
    expect(parseCsv('a,b\n1,2')).toEqual([
      ['a', 'b'],
      ['1', '2']
    ]);
  });

  it('preserves a single non-blank cell on the last unterminated line', () => {
    expect(parseCsv('a,b\nonly-one-cell')).toEqual([
      ['a', 'b'],
      ['only-one-cell']
    ]);
  });
});

describe('toCsv', () => {
  it('joins header and rows with commas and CRLF line endings', () => {
    expect(toCsv(['a', 'b'], [['1', '2'], ['3', '4']])).toBe('a,b\r\n1,2\r\n3,4');
  });

  it('quotes a cell containing a comma', () => {
    expect(toCsv(['name'], [['Svensson, Anna']])).toBe('name\r\n"Svensson, Anna"');
  });

  it('escapes embedded quotes by doubling them, then wraps in quotes', () => {
    expect(toCsv(['note'], [['She said "hi"']])).toBe('note\r\n"She said ""hi"""');
  });

  it('quotes a cell containing a newline', () => {
    expect(toCsv(['note'], [['line1\nline2']])).toBe('note\r\n"line1\nline2"');
  });

  it('round-trips through parseCsv', () => {
    const header: string[] = ['name', 'note'];
    const rows: string[][] = [['Svensson, Anna', 'She said "hi"']];
    expect(parseCsv(toCsv(header, rows))).toEqual([header, ...rows]);
  });
});
