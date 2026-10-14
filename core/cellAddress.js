// ================================================================
// cellAddress.js
// ----------------------------------------------------------------
// PURPOSE : Turn row/column numbers into Excel addresses ("D17", "C:E"),
//           so warnings and screens show exactly what the lead sees in Excel.
//           MIRRORS Excel_Notify\Workbook\CellAddress.cs - keep both in step.
//
// CHANGE LOG
// - 2026-10-13 : Initial version, ported 1:1 from CellAddress.cs.
//     * ES module (import/export), not CommonJS. WHY: the SAME file must run
//       in the browser (admin page on GitHub Pages, <script type="module">,
//       no build step) and in Node (nightly job). Both understand ES modules.
// ================================================================

/** 1 -> "A", 27 -> "AA". */
export function columnLetters(column) {
  let letters = '';
  let n = column;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/** (17, 4) -> "D17". */
export function cellAddressOf(row, column) {
  return columnLetters(column) + row;
}
