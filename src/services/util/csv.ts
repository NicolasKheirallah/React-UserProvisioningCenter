/**
 * Minimal RFC 4180 CSV parser: quoted fields, escaped quotes (""), CR/LF and
 * LF line endings. Returns rows of cells; blank lines are dropped.
 */
export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell: string = '';
  let inQuotes: boolean = false;

  const pushCell = (): void => {
    row.push(cell);
    cell = '';
  };
  const pushRow = (): void => {
    pushCell();
    if (row.length > 1 || (row.length === 1 && row[0].trim() !== '')) {
      rows.push(row);
    }
    row = [];
  };

  for (let i = 0; i < text.length; i++) {
    const ch: string = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cell += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        cell += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      pushCell();
    } else if (ch === '\n') {
      pushRow();
    } else if (ch !== '\r') {
      cell += ch;
    }
  }
  if (cell !== '' || row.length > 0) {
    pushRow();
  }
  return rows;
}

function csvCell(value: string): string {
  return /[",\n\r]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

/** Serializes a header row + data rows to RFC 4180 CSV text (CRLF line endings). */
export function toCsv(header: string[], rows: string[][]): string {
  return [header, ...rows].map((row) => row.map(csvCell).join(',')).join('\r\n');
}

/** Triggers a browser download of `text` as a file named `filename`. */
export function downloadCsv(filename: string, text: string): void {
  const blob: Blob = new Blob([text], { type: 'text/csv;charset=utf-8' });
  const url: string = URL.createObjectURL(blob);
  const anchor: HTMLAnchorElement = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}
