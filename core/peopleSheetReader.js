// ================================================================
// peopleSheetReader.js
// ----------------------------------------------------------------
// PURPOSE : Read an OPTIONAL "People" sheet inside a tracker - the 2nd way to
//           match tracker names to emails (user decision 2026-10-13:
//           shared name map -> tracker People sheet -> Microsoft 365).
//           Only reading (SRP); matching is recipientResolver's job.
//           NEW - Excel_Notify has no equivalent.
//
// FORMAT  : a sheet named People / Staff / Team / Contacts (any case) with a
//           header row (within the first 5 rows) holding a "Name" cell and an
//           "Email" / "Mail" cell; one person per row below it:
//               Name        | Email
//               Abhi        | abhiram.k@mirageaec.com
//               Ruwaisid    | ruwaishid@mirageaec.com
//           Several spellings of the same person = several rows.
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * A non-company or malformed email is SKIPPED with a warning, never
//       used (same @mirageaec.com rule as everywhere else).
//     * LIMIT: the snapshot skips HIDDEN sheets (Excel_Notify rule), so the
//       People sheet must be VISIBLE to be read. A hidden People sheet is
//       reported by the caller, not silently ignored.
//     * No People sheet = empty map, no warning (it is optional).
// ================================================================

import { isAllowedAddress, nameMapKey } from './recipientResolver.js';
import { cellAddressOf } from './cellAddress.js';

const SHEET_NAME = /^\s*(people|staff|team|contacts?)\s*$/i;
const MAX_HEADER_ROW = 5;

/** Name of the People sheet among the given names, or null. */
export function findPeopleSheetName(sheetNames) {
  return sheetNames.find((n) => SHEET_NAME.test(n)) || null;
}

/**
 * @param {{sheets: import('./workbookSnapshot.js').SheetSnapshot[]}} snapshot
 * @returns {{map: Map<string,string>, sheetName: string|null, warnings: {sheetName: string, cell: string, message: string}[]}}
 */
export function readPeopleSheet(snapshot) {
  const result = { map: new Map(), sheetName: null, warnings: [] };
  const sheetName = findPeopleSheetName(snapshot.sheets.map((s) => s.name));
  if (!sheetName) return result;

  const sheet = snapshot.sheets.find((s) => s.name === sheetName);
  result.sheetName = sheetName;

  const header = findHeader(sheet);
  if (!header) {
    result.warnings.push({ sheetName, cell: '-', message: 'No "Name" / "Email" header found in the first 5 rows - sheet not used.' });
    return result;
  }

  for (let row = header.row + 1; row <= sheet.rowCount; row++) {
    const name = sheet.text(row, header.nameColumn);
    const email = sheet.text(row, header.emailColumn);
    if (!name && !email) continue; // blank row
    if (!name) continue;
    if (!isAllowedAddress(email)) {
      result.warnings.push({
        sheetName,
        cell: cellAddressOf(row, header.emailColumn),
        message: `'${email}' is not a valid @mirageaec.com address - "${name}" not used.`,
      });
      continue;
    }
    result.map.set(nameMapKey(name), email.trim());
  }
  return result;
}

/** Header row with a "Name" cell and an "Email"/"Mail" cell; null if none. */
function findHeader(sheet) {
  for (let row = 1; row <= Math.min(MAX_HEADER_ROW, sheet.rowCount); row++) {
    let nameColumn = 0;
    let emailColumn = 0;
    for (let col = 1; col <= sheet.columnCount; col++) {
      const text = sheet.text(row, col).toLowerCase();
      if (!nameColumn && text === 'name') nameColumn = col;
      else if (!emailColumn && /e-?mail|^mail/.test(text)) emailColumn = col;
    }
    if (nameColumn && emailColumn) return { row, nameColumn, emailColumn };
  }
  return null;
}
