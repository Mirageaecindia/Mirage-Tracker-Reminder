// ================================================================
// sheetLayoutDetector.js
// ----------------------------------------------------------------
// PURPOSE : Look at the top rows of ONE sheet and find:
//             - the row with the stage titles ("Engineering Inputs", ...)
//             - every stage block and its Name / Date / Status columns
//             - the Drawing Number / Drawing Name / Remarks columns
//             - which block each role matches by keyword (roleMatcher)
//           Only detection (SRP); reading the rows is trackerReader's job.
//           MIRRORS Excel_Notify\Workbook\SheetLayoutDetector.cs
//                 + Models\SheetLayout.cs + Models\StageBlock.cs
//           - keep all of them in step.
//
// RULES (checked against the real "Project Tracker - MEP.xlsx", 2026-10-12):
//   Title row  : among rows 1-6, the row with the most stage-like titles
//                (a title that matches a role, or has "Name" below it).
//                EW08/F-14/BVLGARI/SE01/SW01 -> row 1.
//   Block      : a non-empty title cell in the title row (not Drawing
//                Number / Drawing Name / Remarks / Review Comments) spans up
//                to 3 columns, ending before the next title.
//                Inside it, Name / Date / Status are found by the sub-header
//                text in the next 3 rows ("Name", "Date", "Status");
//                without sub-headers (SE01 / SW01) they are taken in order:
//                1st = Name, 2nd = Date, 3rd = Status.
//                A title with fewer than 3 columns is not a block.
//   Drawing no.: header containing "drawing number" / "drawing no" / "dwg no";
//                otherwise the column (of the first 4) with the most drawing
//                codes like AMR-EW08-MRG-DWG-GEN-ZZZ-MSM-0302.
//   Drawing name: header "drawing name" / "description"; otherwise the column
//                right of the drawing number.
//   Remarks    : header containing "remark".
//   Not a tracker: no drawing-number column, or no blocks ("Daily status",
//                "HVAC Daily Tracker" have a "Modelors" column but no codes).
//
// CHANGE LOG
// - 2026-10-13 : Initial version, ported 1:1 from SheetLayoutDetector.cs,
//       SheetLayout.cs and StageBlock.cs (same rules, same order, same
//       tie-breaks: "strictly greater" keeps the FIRST best row/column).
//     * Titles found by TEXT, never by fixed letters (BVLGARI is shifted one
//       column: drawing number in A AND B, Drawing Name in C, Engineering
//       Inputs in D:F).
//     * Drawing Number / Name / Remarks headers are searched only in rows
//       1..TitleRow+2 (max 6) - rows 4-6 are already drawing rows; a drawing
//       NAME containing "description" or "remark" must not become a header.
//     * If two blocks match the same role, the FIRST (leftmost) is
//       auto-matched; the lead can change it on the admin page.
//     * C# Dictionary<Role, StageBlock> AutoMatched -> JS Map.
//     * Case-insensitive title compare: C# OrdinalIgnoreCase -> JS
//       toLowerCase() on both sides (same result for these English titles).
//     * Layout is plain data + small getters, so it can be shown on the
//       admin page and compared in the parity check.
// ================================================================

import { matchRole } from './roleMatcher.js';
import { columnLetters } from './cellAddress.js';

const MAX_HEADER_ROW = 6;
const BLOCK_WIDTH = 3;

/** A drawing code: 4+ parts joined by '-', letters/digits only. */
export const DRAWING_CODE = /^[A-Za-z0-9]+(-[A-Za-z0-9]+){3,}$/;

const NOT_A_BLOCK_TITLE = /drawing|dwg|remark|comment|description|s\.?\s*no\b|^no\.?$/i;

// ------------------------------------------------------------
// Data classes (mirror StageBlock.cs / SheetLayout.cs)
// ------------------------------------------------------------

/** ONE stage, e.g. "Engineering Inputs" with its Name / Date / Status columns. */
export class StageBlock {
  constructor(title, nameColumn, dateColumn, statusColumn) {
    /** Header text as typed in the sheet, e.g. "Modeller With Coordination". */
    this.title = title;
    this.nameColumn = nameColumn;
    this.dateColumn = dateColumn;
    this.statusColumn = statusColumn;
  }

  /** "C:E" - shown on screens so the lead can check it. */
  get columnRange() {
    return columnLetters(this.nameColumn) + ':' + columnLetters(this.statusColumn);
  }

  /** "Engineering Inputs (C:E)". */
  toString() {
    return `${this.title} (${this.columnRange})`;
  }
}

/** What the detector found in ONE sheet. */
export class SheetLayout {
  constructor(sheetName) {
    this.sheetName = sheetName;
    /** Row with the stage titles (1-based). 0 = not found. */
    this.titleRow = 0;
    /** 0 = not found. */
    this.drawingNumberColumn = 0;
    this.drawingNameColumn = 0;
    this.remarksColumn = 0;
    /** Every stage block in the header, left to right. */
    this.blocks = [];
    /** Role -> block found by keyword (missing key = keyword not found). */
    this.autoMatched = new Map();
  }

  /** Drawing-number column found AND at least one block (see header). */
  get isTracker() {
    return this.drawingNumberColumn > 0 && this.blocks.length > 0;
  }

  /** Why the sheet is not a tracker (for the admin page); null when it is. */
  get notTrackerReason() {
    if (this.isTracker) return null;
    if (this.drawingNumberColumn === 0) return 'no drawing numbers found';
    return 'no Name / Date / Status columns found';
  }

  /** Block with this exact title (case-insensitive), or null. Title lookup survives inserted columns. */
  findBlock(title) {
    if (title == null) return null;
    const wanted = String(title).toLowerCase();
    return this.blocks.find((b) => b.title.toLowerCase() === wanted) || null;
  }
}

// ------------------------------------------------------------
// Detection
// ------------------------------------------------------------

/**
 * @param {import('./workbookSnapshot.js').SheetSnapshot} sheet
 * @returns {SheetLayout}
 */
export function detectLayout(sheet) {
  const layout = new SheetLayout(sheet.name);
  if (sheet.rowCount === 0) return layout; // empty sheet (C#: Values == null)

  layout.titleRow = findTitleRow(sheet);
  if (layout.titleRow > 0) fillBlocks(sheet, layout);

  // Headers are only searched in the header area, never in drawing rows (see change log).
  const headerEnd = Math.min(MAX_HEADER_ROW, Math.max(layout.titleRow + 2, 3));
  layout.drawingNumberColumn =
    findHeader(sheet, headerEnd, ['drawing number', 'drawing no', 'dwg no', 'dwg. no']) ?? findCodeColumn(sheet);
  layout.drawingNameColumn =
    findHeader(sheet, headerEnd, ['drawing name', 'description', 'drawing title']) ??
    (layout.drawingNumberColumn > 0 ? layout.drawingNumberColumn + 1 : 0);
  layout.remarksColumn = findHeader(sheet, headerEnd, ['remark']) ?? 0;

  for (const block of layout.blocks) {
    const role = matchRole(block.title);
    if (role !== null && !layout.autoMatched.has(role)) layout.autoMatched.set(role, block); // first (leftmost) wins
  }
  return layout;
}

/** Row (1..6) with the most stage-like titles; 0 if none. */
function findTitleRow(sheet) {
  let bestRow = 0;
  let bestScore = 0;
  for (let row = 1; row <= Math.min(MAX_HEADER_ROW, sheet.rowCount); row++) {
    let score = 0;
    for (let col = 1; col <= sheet.columnCount; col++) {
      const text = sheet.text(row, col);
      if (text.length === 0 || NOT_A_BLOCK_TITLE.test(text)) continue;
      if (matchRole(text) !== null || hasSubHeaders(sheet, row, col)) score++;
    }
    if (score > bestScore) {
      bestScore = score;
      bestRow = row;
    }
  }
  return bestRow;
}

/** True if "Name" appears below the title within the next 3 rows (a stage block). */
function hasSubHeaders(sheet, titleRow, col) {
  for (let row = titleRow + 1; row <= titleRow + 3; row++) {
    if (isWord(sheet.text(row, col), 'name')) return true;
  }
  return false;
}

function fillBlocks(sheet, layout) {
  const row = layout.titleRow;
  const titleColumns = [];
  for (let c = 1; c <= sheet.columnCount; c++) {
    if (sheet.text(row, c).length > 0) titleColumns.push(c);
  }

  for (let i = 0; i < titleColumns.length; i++) {
    const start = titleColumns[i];
    const title = sheet.text(row, start);
    if (NOT_A_BLOCK_TITLE.test(title)) continue;

    const nextTitle = i + 1 < titleColumns.length ? titleColumns[i + 1] : Number.MAX_SAFE_INTEGER;
    const end = Math.min(start + BLOCK_WIDTH - 1, nextTitle - 1);
    if (end - start + 1 < BLOCK_WIDTH) continue; // not a Name/Date/Status block (see header)

    layout.blocks.push(createBlock(sheet, row, title, start));
  }
}

/** Block with Name/Date/Status from sub-headers, or in order (see header). */
function createBlock(sheet, titleRow, title, start) {
  let name = null;
  let date = null;
  let status = null;
  for (let row = titleRow + 1; row <= titleRow + 3; row++) {
    for (let col = start; col < start + BLOCK_WIDTH; col++) {
      const text = sheet.text(row, col);
      if (name === null && isWord(text, 'name')) name = col;
      else if (date === null && isWord(text, 'date')) date = col;
      else if (status === null && isWord(text, 'status')) status = col;
    }
  }

  const allFound = name !== null && date !== null && status !== null;
  return allFound
    ? new StageBlock(title, name, date, status)
    : new StageBlock(title, start, start + 1, start + 2);
}

function isWord(text, word) {
  return text != null && String(text).trim().toLowerCase() === word;
}

/**
 * First column (rows 1..lastHeaderRow) whose header contains one of the phrases; null if none.
 * lastHeaderRow keeps the search out of drawing rows (see change log).
 */
function findHeader(sheet, lastHeaderRow, phrases) {
  for (let row = 1; row <= Math.min(lastHeaderRow, sheet.rowCount); row++) {
    for (let col = 1; col <= sheet.columnCount; col++) {
      const text = sheet.text(row, col).toLowerCase();
      if (text.length > 0 && phrases.some((p) => text.includes(p))) return col;
    }
  }
  return null;
}

/** Column (1-4) with the most drawing codes, leftmost on a tie; 0 if none has 3+. */
function findCodeColumn(sheet) {
  let bestCol = 0;
  let bestCount = 2; // need at least 3 codes
  for (let col = 1; col <= Math.min(4, sheet.columnCount); col++) {
    let count = 0;
    for (let row = 1; row <= sheet.rowCount; row++) {
      if (DRAWING_CODE.test(sheet.text(row, col))) count++;
    }
    if (count > bestCount) {
      bestCount = count;
      bestCol = col;
    }
  }
  return bestCol;
}
