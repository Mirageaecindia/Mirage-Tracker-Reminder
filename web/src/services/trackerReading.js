// ================================================================
// web/src/services/trackerReading.js
// ----------------------------------------------------------------
// PURPOSE : Download a tracker AS THE SIGNED-IN ADMIN and run the SAME
//           core reader the nightly job runs (snapshot -> layouts -> read ->
//           People sheet). Shared by the Tracker editor, Name matching and
//           Preview screens (DRY). Only reading (SRP).
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * SheetJS is the browser build of the same version the job uses
//       (0.20.3); it is PASSED IN to core/workbookSnapshot (that file never
//       imports a library itself - see its change log).
//     * The admin's OWN SharePoint access is used here (delegated): a tracker
//       the admin cannot open in SharePoint cannot be read here either.
// ================================================================

import * as XLSX from 'xlsx';
import { downloadTracker } from './trackerLinks.js';
import { takeSnapshot } from '../../../core/workbookSnapshot.js';
import { detectLayout } from '../../../core/sheetLayoutDetector.js';
import { readTracker } from '../../../core/trackerReader.js';
import { readPeopleSheet } from '../../../core/peopleSheetReader.js';
import { toSheetSetups } from '../../../core/trackerConfig.js';

/** Location -> Map(sheetName -> SheetLayout) for the column screens. */
export async function loadWorkbookLayouts(graph, loc) {
  const snapshot = await snapshotOf(graph, loc);
  return new Map(snapshot.sheets.map((s) => [s.name, detectLayout(s)]));
}

/** Registered tracker -> everything the job would read from it. */
export async function readTrackerForPage(graph, tracker) {
  if (!tracker) throw new Error('Choose a tracker first.');
  const snapshot = await snapshotOf(graph, tracker);
  const layouts = new Map(snapshot.sheets.map((s) => [s.name, detectLayout(s)]));
  return {
    snapshot,
    layouts,
    read: readTracker(snapshot, layouts, toSheetSetups(tracker)),
    people: readPeopleSheet(snapshot),
    workbookName: tracker.filePath.split('/').pop(),
  };
}

async function snapshotOf(graph, loc) {
  const bytes = await downloadTracker(graph, loc);
  return takeSnapshot(XLSX, new Uint8Array(bytes), { name: loc.filePath.split('/').pop(), fullName: `${loc.sitePath}/${loc.filePath}` });
}
