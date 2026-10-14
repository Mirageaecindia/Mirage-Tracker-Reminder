// ================================================================
// nameCellParser.js
// ----------------------------------------------------------------
// PURPOSE : Turn ONE Name cell into the people in it:
//             "Abhi"                -> Abhi
//             "Abhi (markup)"       -> Abhi, note "markup"
//             "Akash / Ruwaishid"   -> Akash, Ruwaishid   (both get it)
//             " Ruwaishid "         -> Ruwaishid
//           Only parsing (SRP). Matching to email is recipientResolver's job.
//           MIRRORS Excel_Notify\Workbook\NameCellParser.cs - keep both in step.
//
// RULES (same as C#):
//     * Separators: / , & ; + and the word "and".
//     * Text in (...) or [...] is a NOTE, not part of the name, and applies
//       to the whole cell ("Abhi (markup)", "Ruwaishid (WFH)").
//     * Repeated spaces collapsed; empty parts dropped.
//     * A checkbox (true/false) in a Name column is NOT a name.
//     * Spelling is NOT corrected ("Ruwaisid" stays "Ruwaisid") - the lead
//       maps it once and the shared name map remembers it.
//
// CHANGE LOG
// - 2026-10-13 : Initial version, ported 1:1 from NameCellParser.cs.
//     * C# Regex.Split -> String.split(regex): same result here because the
//       separator pattern has no capturing group.
//     * BRACKET has the /g flag in JS (needed by matchAll / replace-all);
//       C# Matches / Replace are "all" by default.
// ================================================================

const BRACKET = /[([]([^)\]]*)[)\]]/g;
const SEPARATOR = /\s*(?:\/|,|&|;|\+|\band\b)\s*/i;
const SPACES = /\s+/g;

/**
 * People in the cell; empty list for an empty cell.
 * @param {*} cell  raw cell value from the snapshot
 * @returns {{name: string, note: string}[]}
 */
export function parseNameCell(cell) {
  if (cell == null || typeof cell === 'boolean') return [];
  const text = typeof cell === 'string' ? cell : String(cell);
  if (text.trim() === '') return [];

  // Notes in brackets, then remove them from the names.
  const note = [...text.matchAll(BRACKET)]
    .map((m) => m[1].trim())
    .filter((n) => n.length > 0)
    .join(', ');
  const names = text.replace(BRACKET, ' ');

  return names
    .split(SEPARATOR)
    .map((n) => n.replace(SPACES, ' ').trim())
    .filter((n) => n.length > 0)
    .map((n) => ({ name: n, note }));
}
