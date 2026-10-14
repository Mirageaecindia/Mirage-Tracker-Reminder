// ================================================================
// dateCellParser.js
// ----------------------------------------------------------------
// PURPOSE : Turn ONE Date cell into a due date, or "completed", or a
//           problem to warn about. Only parsing (SRP).
//           MIRRORS Excel_Notify\Workbook\DateCellParser.cs - keep both in step.
//
// RULES (user decisions 2026-10-12, same as C#):
//           real date                 -> due date
//           "COMPLETED" / "COMPLETE"  -> item is DONE
//           empty                     -> no due date (no warning)
//           any other text            -> no due date + WARNING (cell address)
//           checkbox in a Date column -> WARNING (not a date)
//
// DATE FORMAT IN JS: dates are plain "YYYY-MM-DD" strings, never JS Date
//           objects outside this file. WHY: a JS Date carries a time zone;
//           the nightly job runs in UTC on GitHub but "today" means India
//           time. A date-only string cannot shift by a day.
//
// CHANGE LOG
// - 2026-10-13 : Initial version, ported 1:1 from DateCellParser.cs.
//     * Typed dates: DAY-FIRST formats only, exactly the C# list:
//         dd-MM-yyyy d-M-yyyy dd/MM/yyyy d/M/yyyy dd.MM.yyyy d.M.yyyy
//         dd-MM-yy   d-M-yy   dd/MM/yy   d/M/yy
//         dd-MMM-yyyy d-MMM-yyyy "dd MMM yyyy" "d MMM yyyy"
//       WHY: 08-10-2026 = 8 October; month-first would silently swap.
//       C# DateTime.TryParseExact is not available in JS, so the formats are
//       written out as regexes below (TYPED_* ) - same accepted inputs:
//         - day/month 1 or 2 digits ("d"/"dd", "M"/"MM")
//         - 2-digit years only with "-" or "/" (the C# list has no "dd.MM.yy")
//         - 2-digit year 00-49 -> 20xx, 50-99 -> 19xx (InvariantCulture's
//           TwoDigitYearMax = 2049)
//         - month names Jan..Dec, case-insensitive, 4-digit year only
//         - impossible dates (31-02-2026) are rejected, like C#.
//     * A number (date stored as a plain number in a General cell) is read
//       as an Excel serial date only between 2000 and 2100 (C# FromOADate
//       rule); anything else warns.
//     * "06/102026" (seen in EW08) matches no format -> warning, as agreed.
// ================================================================

export const DateCellKind = Object.freeze({
  Empty: 'Empty',
  Date: 'Date',
  Completed: 'Completed',
  Invalid: 'Invalid',
});

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

// d-M-yyyy / d/M/yyyy / d.M.yyyy (separator must be the same twice: \2)
const TYPED_NUMERIC_4 = /^(\d{1,2})([-/.])(\d{1,2})\2(\d{4})$/;
// d-M-yy / d/M/yy (no dot - matches the C# list)
const TYPED_NUMERIC_2 = /^(\d{1,2})([-/])(\d{1,2})\2(\d{2})$/;
// d-MMM-yyyy / d MMM yyyy
const TYPED_MONTH_NAME = /^(\d{1,2})([- ])([A-Za-z]{3})\2(\d{4})$/;

/**
 * @param {*} cell  raw value from the snapshot: null, string, number, boolean or Date
 * @returns {{kind: string, date: string|null, raw: string|null}}
 */
export function parseDateCell(cell) {
  if (cell == null) return result(DateCellKind.Empty);
  if (cell instanceof Date) return result(DateCellKind.Date, isoFromUtcDate(cell));
  if (typeof cell === 'number') return fromSerial(cell);
  if (typeof cell === 'boolean') return result(DateCellKind.Invalid, null, cell ? 'True' : 'False');

  const text = String(cell).trim();
  if (text.length === 0) return result(DateCellKind.Empty);

  const upper = text.toUpperCase();
  if (upper === 'COMPLETED' || upper === 'COMPLETE') return result(DateCellKind.Completed, null, text);

  const typed = parseTyped(text);
  if (typed) return result(DateCellKind.Date, typed, text);

  return result(DateCellKind.Invalid, null, text);
}

/** Excel serial number -> date, only 2000..2100 (see change log). */
function fromSerial(number) {
  // C# DateTime.FromOADate: day 0 = 1899-12-30, fraction = time of day.
  if (Number.isFinite(number)) {
    const ms = Date.UTC(1899, 11, 30) + Math.floor(number) * 86400000;
    const date = new Date(ms);
    const year = date.getUTCFullYear();
    if (year >= 2000 && year <= 2100) return result(DateCellKind.Date, isoFromUtcDate(date));
  }
  return result(DateCellKind.Invalid, null, String(number));
}

/** Day-first typed text -> "YYYY-MM-DD", or null when no format matches. */
function parseTyped(text) {
  let m = TYPED_NUMERIC_4.exec(text);
  if (m) return makeIso(+m[4], +m[3], +m[1]);

  m = TYPED_NUMERIC_2.exec(text);
  if (m) {
    const yy = +m[4];
    return makeIso(yy <= 49 ? 2000 + yy : 1900 + yy, +m[3], +m[1]); // TwoDigitYearMax 2049
  }

  m = TYPED_MONTH_NAME.exec(text);
  if (m) {
    const month = MONTHS.indexOf(m[3].toLowerCase()) + 1;
    if (month > 0) return makeIso(+m[4], month, +m[1]);
  }
  return null;
}

/** Valid calendar date -> "YYYY-MM-DD"; null for 31-02 etc. */
function makeIso(year, month, day) {
  if (month < 1 || month > 12 || day < 1) return null;
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null; // rolled over = impossible date
  return isoFromUtcDate(date);
}

/** UTC parts of a Date -> "YYYY-MM-DD" (time dropped, like C# .Date). */
export function isoFromUtcDate(date) {
  const y = String(date.getUTCFullYear()).padStart(4, '0');
  const m = String(date.getUTCMonth() + 1).padStart(2, '0');
  const d = String(date.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function result(kind, date = null, raw = null) {
  return { kind, date, raw };
}
