// ================================================================
// trackerReader.js
// ----------------------------------------------------------------
// PURPOSE : Read the drawing rows of the chosen sheets into AssignmentItems,
//           using the layout found by sheetLayoutDetector and the lead's
//           SheetSetup (roles + blocks). Works on the snapshot only. Only
//           reading (SRP); filtering / grouping is recipientBuilder's job.
//           MIRRORS Excel_Notify\Workbook\TrackerReader.cs + Models\ReadResult.cs
//           + Models\ReadWarning.cs - keep them in step.
//
// ROW TYPES (walking down from the row below the stage titles):
//   Drawing row   : Drawing Number is a drawing code, OR the Drawing Name is
//                   filled and a chosen role has a name (SE01 row 7 has a
//                   name but no number).
//   Sub-header row: "Name" / "Date" / "Status" in a block -> label row only.
//   Label row     : text in the columns up to Drawing Name:
//                     known discipline ("MECHANICAL", "PLUMBING AND DRAINAGE",
//                     "FIRE PROTECTION", ...) -> current discipline
//                     anything else ("AIR CONDITIONING SYSTEM LAYOUT")
//                       -> current sub-group
//   Blank row     : ignored.
//
// PER DRAWING ROW, PER CHOSEN ROLE:
//   Name cell   -> 0..n people (nameCellParser). Empty -> counted as unassigned.
//   Date cell   -> due date / COMPLETED (= done) / bad text (= warning)
//   Status cell -> ticked = done
//   One AssignmentItem per person.
//
// CHANGE LOG
// - 2026-10-13 : Initial version, ported 1:1 from TrackerReader.cs.
//     * Disciplines are an EXACT list, not "contains" - "FIRE FIGHTING
//       LAYOUT" (EW08 row 216) is a sub-group, not a discipline.
//     * Consecutive label rows are joined ("APPARTMENT / AIR CONDITIONING
//       SYSTEM LAYOUT", EW08 rows 96-97).
//     * Font colour (red text in EW08) is ignored (user decision).
//     * Status accepts TRUE (checkbox) and the texts TRUE / YES / DONE.
//     * A label row with several texts keeps ALL of them joined with " / "
//       (BVLGARI row 34: "Mansion A1 / AIR CONDITIONING SYSTEM LAYOUT").
//     * C# ReadResult / ReadWarning classes -> plain objects here
//       (createReadResult / warning text built the same way).
//     * Roles of a drawing row are read in role order (C# OrderBy(p.Key)).
// ================================================================

import { ALL_ROLES, roleShortName } from './roles.js';
import { cellAddressOf } from './cellAddress.js';
import { parseNameCell } from './nameCellParser.js';
import { parseDateCell, DateCellKind } from './dateCellParser.js';
import { DRAWING_CODE } from './sheetLayoutDetector.js';
import { AssignmentItem } from './assignmentItem.js';

// Exact discipline names, compared after normalise() (see change log).
const DISCIPLINES = new Set([
  'MECHANICAL', 'HVAC', 'ELECTRICAL', 'ELV', 'LOW CURRENT',
  'PLUMBING', 'DRAINAGE', 'PLUMBING AND DRAINAGE', 'PHE', 'PUBLIC HEALTH',
  'FIRE PROTECTION', 'FIRE FIGHTING', 'FIRE', 'FIRE ALARM',
]);

const SPACES = /\s+/g;

/** Everything read from the chosen sheets (mirrors ReadResult.cs). */
export function createReadResult() {
  return {
    /** @type {AssignmentItem[]} */
    items: [],
    /** @type {{sheetName: string, cell: string, message: string}[]} */
    warnings: [],
    /** sheet name -> SheetSetup that was read */
    setups: new Map(),
    /** Drawing rows with an empty Name cell in a chosen role column (counted, never emailed). */
    unassignedCount: 0,
  };
}

/** "EW08 (100%)!D17 - '06/102026' is not a date..." (mirrors ReadWarning.ToString). */
export function warningText(w) {
  return `${w.sheetName}!${w.cell} - ${w.message}`;
}

/**
 * Reads every setup's sheet. Sheets / blocks that are missing become warnings.
 * @param {{sheets: import('./workbookSnapshot.js').SheetSnapshot[]}} snapshot
 * @param {Map<string, import('./sheetLayoutDetector.js').SheetLayout>} layouts  sheet name -> layout
 * @param {Iterable<import('./assignmentItem.js').SheetSetup>} setups
 */
export function readTracker(snapshot, layouts, setups) {
  const result = createReadResult();
  for (const setup of setups) {
    result.setups.set(setup.sheetName, setup);

    const sheet = snapshot.sheets.find((s) => s.name === setup.sheetName);
    const layout = layouts.get(setup.sheetName);
    if (!sheet || !layout) {
      warn(result, setup.sheetName, '-', 'Sheet not found - skipped.');
      continue;
    }

    const blocks = new Map();
    for (const role of setup.roles) {
      const title = setup.blockTitles.get(role);
      const block = title !== undefined ? layout.findBlock(title) : null;
      if (!block) warn(result, sheet.name, '-', `${roleShortName(role)}: column "${title ?? ''}" not found - skipped.`);
      else blocks.set(role, block);
    }

    if (blocks.size > 0) readSheet(sheet, layout, blocks, result);
  }
  return result;
}

function readSheet(sheet, layout, blocks, result) {
  let discipline = '';
  let subGroup = '';
  let previousWasLabel = false;
  const lastLabelColumn = Math.max(layout.drawingNumberColumn, layout.drawingNameColumn);

  for (let row = layout.titleRow + 1; row <= sheet.rowCount; row++) {
    const number = sheet.text(row, layout.drawingNumberColumn);
    const name = layout.drawingNameColumn > 0 ? sheet.text(row, layout.drawingNameColumn) : '';
    const subHeader = isSubHeaderRow(sheet, row, blocks.values());
    const isCode = DRAWING_CODE.test(number);

    if (!subHeader && (isCode || (name.length > 0 && anyName(sheet, row, blocks.values())))) {
      readDrawingRow(sheet, layout, blocks, row, isCode ? number : '', name, discipline, subGroup, result);
      previousWasLabel = false;
      continue;
    }

    // Label row (discipline / sub-group) or blank row.
    const labels = labelTexts(sheet, row, lastLabelColumn);
    if (labels.length === 0) continue; // blank: keeps the current headings

    const disciplineLabel = labels.find((l) => DISCIPLINES.has(normalise(l)));
    if (disciplineLabel !== undefined) {
      discipline = disciplineLabel.replace(SPACES, ' ').trim();
      subGroup = '';
    } else {
      // All texts of the row ("Mansion A1 / AIR CONDITIONING SYSTEM LAYOUT", BVLGARI row 34).
      const label = labels.map((l) => l.replace(SPACES, ' ').trim()).join(' / ');
      subGroup = previousWasLabel && subGroup.length > 0 ? `${subGroup} / ${label}` : label;
    }
    previousWasLabel = true;
  }
}

function readDrawingRow(sheet, layout, blocks, row, number, name, discipline, subGroup, result) {
  const remarks = layout.remarksColumn > 0 ? sheet.text(row, layout.remarksColumn) : '';

  // Role order, like C# blocks.OrderBy(p => p.Key).
  for (const role of ALL_ROLES) {
    const block = blocks.get(role);
    if (!block) continue;

    const people = parseNameCell(sheet.get(row, block.nameColumn));
    if (people.length === 0) {
      result.unassignedCount++;
      continue;
    }

    const date = parseDateCell(sheet.get(row, block.dateColumn));
    if (date.kind === DateCellKind.Invalid) {
      warn(result, sheet.name, cellAddressOf(row, block.dateColumn),
        `'${date.raw}' in ${block.title} Date is not a date - sent with no due date.`);
    }

    const done = isTicked(sheet.get(row, block.statusColumn)) || date.kind === DateCellKind.Completed;

    for (const person of people) {
      result.items.push(new AssignmentItem({
        sheetName: sheet.name,
        discipline,
        subGroup,
        drawingNumber: number,
        drawingName: name,
        role,
        stageTitle: block.title,
        personName: person.name,
        note: person.note,
        dueDate: date.date,
        isDone: done,
        remarks,
        cellAddress: cellAddressOf(row, block.nameColumn),
      }));
    }
  }
}

// ------------------------------------------------------------
// Helpers
// ------------------------------------------------------------

/** True if a block's cells hold "Name" / "Date" / "Status" (the sub-header row). */
function isSubHeaderRow(sheet, row, blocks) {
  for (const b of blocks) {
    if (isWord(sheet.text(row, b.nameColumn), 'name')
      || isWord(sheet.text(row, b.dateColumn), 'date')
      || isWord(sheet.text(row, b.statusColumn), 'status')) return true;
  }
  return false;
}

function anyName(sheet, row, blocks) {
  for (const b of blocks) {
    if (parseNameCell(sheet.get(row, b.nameColumn)).length > 0) return true;
  }
  return false;
}

/** Text cells in columns 1..lastColumn (no drawing codes, no TRUE/FALSE). */
function labelTexts(sheet, row, lastColumn) {
  const labels = [];
  for (let col = 1; col <= Math.max(1, lastColumn); col++) {
    if (typeof sheet.get(row, col) !== 'string') continue;
    const text = sheet.text(row, col);
    if (text.length > 0 && !DRAWING_CODE.test(text) && !labels.includes(text)) labels.push(text);
  }
  return labels;
}

function isTicked(cell) {
  if (typeof cell === 'boolean') return cell;
  if (cell == null) return false;
  const text = String(cell).trim().toUpperCase();
  return text === 'TRUE' || text === 'YES' || text === 'DONE';
}

/** "Plumbing  & Drainage" -> "PLUMBING AND DRAINAGE". */
function normalise(text) {
  return text.replace(/&/g, ' AND ').replace(SPACES, ' ').trim().toUpperCase();
}

function isWord(text, word) {
  return text != null && String(text).trim().toLowerCase() === word;
}

function warn(result, sheetName, cell, message) {
  result.warnings.push({ sheetName, cell, message });
}
