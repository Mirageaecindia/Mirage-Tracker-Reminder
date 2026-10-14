// ================================================================
// parityCheck.js
// ----------------------------------------------------------------
// PURPOSE : Prove the JS port reads a tracker EXACTLY like Excel_Notify.
//           Reads one .xlsx and prints what Excel_Notify's windows show,
//           so the two can be compared side by side:
//             1. Sheets list   (= Excel_Notify Setup window)
//                "EW08 (100%) - Engineers C:E, Modelers F:H, ..."
//             2. Per sheet x role counts (items / done / pending / overdue)
//             3. People grid   (= Notify window with Show = Pending, ALL 4
//                roles ticked): name | Eng | Mod | Ann | QA | overdue | earliest due
//             4. Warnings      (= Warnings window)
//           Read-only: never changes the tracker.
//
// USAGE   : node tools/parityCheck.js "<path to tracker.xlsx>" [--today YYYY-MM-DD]
//           --today defaults to the current date in India time.
//
// COMPARE WITH EXCEL_NOTIFY: open the same file in Excel -> Notify Assignees
//           -> tick EVERY tracker sheet and ALL 4 roles (accept the auto-
//           matched columns) -> Notify window, Show = Pending.
//
// CHANGE LOG
// - 2026-10-13 : Initial version.
//     * Uses each sheet's AUTO-MATCHED columns for all 4 roles (no lead
//       picks). Roles without a keyword match (SE01/SW01 "Inputs") are
//       listed as "not matched" - in Excel_Notify the lead would pick them.
//     * People are grouped by tracker name (case/space-insensitive) per
//       sheet. Excel_Notify additionally MERGES spellings that resolve to the
//       same email - that needs the company directory, which this offline
//       check does not have. Merged rows there = several rows here.
// ================================================================

import { readFileSync } from 'node:fs';
import { basename } from 'node:path';
import XLSX from 'xlsx';

import { ALL_ROLES, roleShortName } from '../core/roles.js';
import { takeSnapshot } from '../core/workbookSnapshot.js';
import { detectLayout } from '../core/sheetLayoutDetector.js';
import { SheetSetup } from '../core/assignmentItem.js';
import { readTracker, warningText } from '../core/trackerReader.js';

const args = process.argv.slice(2);
const file = args.find((a) => !a.startsWith('--'));
if (!file) {
  console.error('Usage: node tools/parityCheck.js "<tracker.xlsx>" [--today YYYY-MM-DD]');
  process.exit(1);
}
const todayArg = args.indexOf('--today');
const today = todayArg >= 0 ? args[todayArg + 1] : indiaToday();

const snapshot = takeSnapshot(XLSX, readFileSync(file), { name: basename(file), fullName: file });
const layouts = new Map(snapshot.sheets.map((s) => [s.name, detectLayout(s)]));

// ---- 1. Sheets list ---------------------------------------------------------
section(`1. SHEETS  (${snapshot.name}, today = ${today})`);
for (const sheet of snapshot.sheets) {
  const layout = layouts.get(sheet.name);
  if (!layout.isTracker) {
    console.log(`  ${sheet.name} - not a tracker (${layout.notTrackerReason})`);
    continue;
  }
  const matched = ALL_ROLES.map((r) => {
    const b = layout.autoMatched.get(r);
    return b ? `${roleShortName(r)} ${b.columnRange}` : `${roleShortName(r)} (not matched)`;
  });
  console.log(`  ${sheet.name} - ${matched.join(', ')}`);
  console.log(`      title row ${layout.titleRow} | drawing no. col ${layout.drawingNumberColumn}`
    + ` | name col ${layout.drawingNameColumn} | remarks col ${layout.remarksColumn}`
    + ` | blocks: ${layout.blocks.map(String).join('; ')}`);
}
if (snapshot.hiddenSheets.length) console.log(`  (hidden, skipped: ${snapshot.hiddenSheets.join(', ')})`);

// ---- read with all auto-matched roles ----------------------------------------
const setups = [];
for (const [name, layout] of layouts) {
  if (!layout.isTracker) continue;
  const setup = new SheetSetup(name);
  for (const [role, block] of layout.autoMatched) {
    setup.roles.add(role);
    setup.blockTitles.set(role, block.title);
  }
  if (setup.roles.size > 0) setups.push(setup);
}
const result = readTracker(snapshot, layouts, setups);

// ---- 2. Per sheet x role counts -------------------------------------------------
section('2. COUNTS per sheet x role  (items = one person x drawing)');
console.log(pad('Sheet', 26) + pad('Role', 12) + pad('Items', 7) + pad('Done', 6) + pad('Pending', 9) + 'Overdue');
for (const setup of setups) {
  for (const role of ALL_ROLES) {
    if (!setup.roles.has(role)) continue;
    const items = result.items.filter((i) => i.sheetName === setup.sheetName && i.role === role);
    console.log(pad(setup.sheetName, 26) + pad(roleShortName(role), 12) + pad(items.length, 7)
      + pad(items.filter((i) => i.isDone).length, 6) + pad(items.filter((i) => !i.isDone).length, 9)
      + items.filter((i) => i.isOverdue(today)).length);
  }
}
console.log(`\n  Total items: ${result.items.length} | Unassigned (empty Name cells): ${result.unassignedCount}`);

// ---- 3. People grid (Show = Pending) -------------------------------------------
section('3. PEOPLE  (Show = Pending; counts are PENDING items per role)');
console.log(pad('Sheet', 26) + pad('Name (tracker)', 22) + pad('Eng', 5) + pad('Mod', 5) + pad('Ann', 5) + pad('QA', 5)
  + pad('Overdue', 9) + 'Earliest due');
const rows = new Map();
for (const item of result.items.filter((i) => !i.isDone)) {
  const key = `${item.sheetName}|${nameKey(item.personName)}`;
  if (!rows.has(key)) rows.set(key, { sheet: item.sheetName, name: item.personName.trim(), items: [] });
  rows.get(key).items.push(item);
}
[...rows.values()]
  .sort((a, b) => a.sheet.localeCompare(b.sheet) || a.name.localeCompare(b.name))
  .forEach((r) => {
    const count = (role) => r.items.filter((i) => i.role === role).length;
    const dues = r.items.map((i) => i.dueDate).filter(Boolean).sort();
    console.log(pad(r.sheet, 26) + pad(r.name, 22) + ALL_ROLES.map((role) => pad(count(role), 5)).join('')
      + pad(r.items.filter((i) => i.isOverdue(today)).length, 9) + (dues[0] ? dayFirst(dues[0]) : '-'));
  });

// ---- 4. Warnings ----------------------------------------------------------------
section(`4. WARNINGS (${result.warnings.length})`);
result.warnings.forEach((w) => console.log('  ' + warningText(w)));

// ------------------------------------------------------------
function section(title) {
  console.log('\n' + '='.repeat(78) + '\n' + title + '\n' + '='.repeat(78));
}
function pad(value, width) {
  const text = String(value);
  return text.length >= width ? text.slice(0, width - 1) + ' ' : text + ' '.repeat(width - text.length);
}
function nameKey(name) {
  return name.trim().split(/\s+/).join(' ').toLowerCase();
}
function dayFirst(iso) {
  const [y, m, d] = iso.split('-');
  return `${d}-${m}-${y}`;
}
/** Today's date in India (UTC+05:30) as YYYY-MM-DD - the job server runs in UTC. */
function indiaToday() {
  const ist = new Date(Date.now() + 5.5 * 3600 * 1000);
  return ist.toISOString().slice(0, 10);
}
