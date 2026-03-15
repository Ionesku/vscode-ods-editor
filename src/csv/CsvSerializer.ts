import { SpreadsheetModel } from '../model/SpreadsheetModel';

/** Serialize the active sheet of a SpreadsheetModel to a CSV string */
export function modelToCsv(model: SpreadsheetModel): string {
  const sheet = model.activeSheet;
  const usedRange = sheet.getUsedRange();
  if (!usedRange) return '';

  const rows: string[] = [];
  for (let r = usedRange.startRow; r <= usedRange.endRow; r++) {
    const cells: string[] = [];
    for (let c = usedRange.startCol; c <= usedRange.endCol; c++) {
      const cell = sheet.getCell(c, r);
      const value = cell.computedValue ?? cell.rawValue;
      cells.push(csvEscape(value));
    }
    rows.push(cells.join(','));
  }
  return rows.join('\r\n');
}

/** Parse a CSV string into a SpreadsheetModel */
export function csvToModel(csv: string, sheetName = 'Sheet1'): SpreadsheetModel {
  const model = new SpreadsheetModel();
  model.activeSheet.name = sheetName;
  const sheet = model.activeSheet;

  const lines = csv.split(/\r?\n/);
  for (let r = 0; r < lines.length; r++) {
    const line = lines[r];
    if (!line.trim()) continue;
    const cells = parseCsvRow(line);
    for (let c = 0; c < cells.length; c++) {
      const raw = cells[c];
      if (raw === '') continue;
      const num = Number(raw);
      const value = !isNaN(num) && raw !== '' ? num : raw;
      sheet.setCell(c, r, {
        rawValue: value,
        formula: null,
        computedValue: value,
        styleId: null,
        mergeColSpan: 1,
        mergeRowSpan: 1,
        mergedInto: null,
      });
    }
  }
  return model;
}

function csvEscape(value: unknown): string {
  if (value === null || value === undefined) return '';
  const s = String(value);
  if (s.includes(',') || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function parseCsvRow(line: string): string[] {
  const result: string[] = [];
  let i = 0;
  while (i <= line.length) {
    if (line[i] === '"') {
      // Quoted field
      let field = '';
      i++; // skip opening quote
      while (i < line.length) {
        if (line[i] === '"' && line[i + 1] === '"') {
          field += '"';
          i += 2;
        } else if (line[i] === '"') {
          i++; // skip closing quote
          break;
        } else {
          field += line[i++];
        }
      }
      result.push(field);
      if (line[i] === ',') i++; // skip comma
    } else {
      // Unquoted field
      const start = i;
      while (i < line.length && line[i] !== ',') i++;
      result.push(line.slice(start, i));
      if (line[i] === ',') i++;
    }
    if (i > line.length) break;
  }
  return result;
}
