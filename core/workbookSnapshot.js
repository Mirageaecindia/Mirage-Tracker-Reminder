// ================================================================
// workbookSnapshot.js
// ----------------------------------------------------------------
// PURPOSE : Plain copy of a workbook's cell VALUES, read from an .xlsx file.
//           Everything after this works on the copy, never on the file
//           library. Only reading (SRP).
//           REPLACES Excel_Notify\Models\WorkbookSnapshot.cs +
//           Workbook\WorkbookSnapshotter.cs (Excel COM -> SheetJS).
//
// GOAL    : give the detector / reader EXACTLY what Excel's Range.Value gave
//           the C# add-in, so the ported logic behaves the same:
//             text cell            -> string
//             number (General)     -> number
//             date-formatted cell  -> Date (UTC midnight)   [C#: DateTime]
//             checkbox / TRUE      -> boolean
//             empty                -> null
//             merged cells         -> value only in the top-left cell (as Excel)
//           Rows / columns are 1-based like Excel, so cell addresses in
//           warnings match what the lead sees.
//
// CHANGE LOG
// - 2026-10-13 : Initial version.
//     * The SheetJS library is PASSED IN (takeSnapshot(XLSX, ...)), not
//       imported here. WHY (Dependency Inversion): the browser loads SheetJS
//       from a <script> tag, Node from node_modules - this file works with
//       either and stays testable.
//     * Read with cellDates:false + cellNF:true, then a number is turned into
//       a Date only when its number FORMAT is a date format (SSF.is_date).
//       WHY: that is exactly how Excel decides Range.Value = DateTime; and
//       SheetJS's own cellDates option builds Dates in the PC's local time
//       zone, which can shift a due date by one day.
//     * 1904 date system workbooks (old Mac files) are converted too.
//     * Hidden / very hidden sheets are skipped and listed (same rule as C#:
//       "EW08 (old)" is hidden and must not be read).
//     * Last row / column = last cell that holds a value or a formula, NOT
//       the sheet's stored dimension. WHY: same reason C# used Find("*")
//       instead of UsedRange - formatting-only cells can make the stored
//       range huge.
//     * An error cell (#N/A, #REF!) becomes its error TEXT. Small difference
//       from C# (Excel returned an error number); both end up as "not a date"
//       / not a drawing, so behaviour is the same.
//     * A sheet that fails to read is skipped, not fatal (same as C#).
// ================================================================

/** One worksheet's values. */
export class SheetSnapshot {
  constructor(name) {
    this.name = name;
    /** Last row / column holding data (1-based). 0 = empty sheet. */
    this.rowCount = 0;
    this.columnCount = 0;
    /** values[row][column], 1-based; null = empty. */
    this.values = [];
  }

  /** Value of one cell; null outside the copied range. */
  get(row, column) {
    if (row < 1 || column < 1 || row > this.rowCount || column > this.columnCount) return null;
    const value = this.values[row][column];
    return value === undefined ? null : value;
  }

  /**
   * Cell text, trimmed; empty string for null.
   * Matches C# value.ToString().Trim() closely enough for the detector:
   * booleans -> "True"/"False" (C# casing), dates -> "dd-MM-yyyy".
   */
  text(row, column) {
    const value = this.get(row, column);
    if (value == null) return '';
    if (typeof value === 'boolean') return value ? 'True' : 'False';
    if (value instanceof Date) return formatDayFirst(value);
    return String(value).trim();
  }
}

/**
 * Reads every visible sheet of an .xlsx file.
 * @param {object} XLSX         the SheetJS library object
 * @param {ArrayBuffer|Uint8Array|Buffer} data  file bytes
 * @param {{name?: string, fullName?: string}} info  file name / SharePoint URL
 * @returns {{name: string, fullName: string, sheets: SheetSnapshot[], hiddenSheets: string[]}}
 */
export function takeSnapshot(XLSX, data, info = {}) {
  const isNodeBuffer = typeof Buffer !== 'undefined' && Buffer.isBuffer(data);
  const workbook = XLSX.read(data, {
    type: isNodeBuffer ? 'buffer' : 'array',
    cellDates: false, // see change log
    cellNF: true,     // number formats -> tell dates from numbers
    cellStyles: false,
    sheetStubs: false,
  });

  const snapshot = {
    name: info.name || 'Workbook',
    fullName: info.fullName || info.name || 'Workbook',
    sheets: [],
    hiddenSheets: [],
  };
  const date1904 = Boolean(workbook.Workbook && workbook.Workbook.WBProps && workbook.Workbook.WBProps.date1904);
  const sheetProps = (workbook.Workbook && workbook.Workbook.Sheets) || [];

  workbook.SheetNames.forEach((name, index) => {
    const hidden = sheetProps[index] && sheetProps[index].Hidden; // 1 = hidden, 2 = very hidden
    if (hidden) {
      snapshot.hiddenSheets.push(name);
      return;
    }
    try {
      snapshot.sheets.push(readSheet(XLSX, workbook.Sheets[name], name, date1904));
    } catch (e) {
      // one unreadable sheet must not stop the others (see change log)
    }
  });
  return snapshot;
}

// ------------------------------------------------------------
// Helpers
// ------------------------------------------------------------

function readSheet(XLSX, sheet, name, date1904) {
  const result = new SheetSnapshot(name);
  const cells = [];
  let lastRow = 0;
  let lastColumn = 0;

  for (const address of Object.keys(sheet)) {
    if (address[0] === '!') continue; // "!ref", "!merges", ... are not cells
    const cell = sheet[address];
    if (cell.v === undefined && !cell.f) continue; // formatting only (see change log)

    const pos = XLSX.utils.decode_cell(address); // 0-based
    const row = pos.r + 1;
    const column = pos.c + 1;
    cells.push([row, column, cellValue(XLSX, cell, date1904)]);
    if (row > lastRow) lastRow = row;
    if (column > lastColumn) lastColumn = column;
  }

  if (lastRow === 0 || lastColumn === 0) return result; // empty sheet

  result.rowCount = lastRow;
  result.columnCount = lastColumn;
  result.values = Array.from({ length: lastRow + 1 }, () => new Array(lastColumn + 1).fill(null));
  for (const [row, column, value] of cells) result.values[row][column] = value;
  return result;
}

/** SheetJS cell -> the value Excel's Range.Value would give (see GOAL). */
function cellValue(XLSX, cell, date1904) {
  switch (cell.t) {
    case 'b':
      return Boolean(cell.v);
    case 'n':
      if (cell.z && XLSX.SSF.is_date(cell.z)) return serialToUtcDate(cell.v, date1904);
      return cell.v;
    case 'd':
      return cell.v instanceof Date ? cell.v : null;
    case 'e':
      return cell.w != null ? String(cell.w) : null; // see change log
    case 's':
      return cell.v == null ? null : String(cell.v);
    default:
      return cell.v == null ? null : cell.v;
  }
}

/** Excel serial -> Date at UTC midnight (day 0 = 1899-12-30; 1904 system shifted by 1462 days). */
function serialToUtcDate(serial, date1904) {
  const days = Math.floor(serial) + (date1904 ? 1462 : 0);
  return new Date(Date.UTC(1899, 11, 30) + days * 86400000);
}

function formatDayFirst(date) {
  const d = String(date.getUTCDate()).padStart(2, '0');
  const m = String(date.getUTCMonth() + 1).padStart(2, '0');
  return `${d}-${m}-${date.getUTCFullYear()}`;
}
