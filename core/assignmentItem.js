// ================================================================
// assignmentItem.js
// ----------------------------------------------------------------
// PURPOSE : Plain data for the reader (SRP, no logic beyond small getters):
//             - AssignmentItem : ONE person, ONE drawing, ONE role
//               ("Vasanth - Modelers - AMR-EW08-...-1001, due 12-10-2026")
//             - SheetSetup     : the lead's choice for ONE sheet - which
//               roles to read and which stage block (by TITLE) each uses
//           MIRRORS Excel_Notify\Models\AssignmentItem.cs + Models\SheetSetup.cs
//           - keep them in step.
//
// RULES (same as C#):
//     * A cell "Akash / Ruwaishid" gives TWO items (both get it).
//     * dueDate = the Date column; "YYYY-MM-DD" string or null.
//     * isDone = Status ticked OR Date says "COMPLETED".
//     * Overdue = not done AND due date before today.
//     * key = what the sent log remembers (with the email): sheet + drawing
//       + role + due date. A changed due date or person = a new email.
//     * Drawing WITHOUT a number: key uses drawing name + cell address
//       (BVLGARI rows 43-45 have no number and the SAME name).
//     * SheetSetup stores the block by TITLE, not letters: if someone inserts
//       a column, the title is found again; a vanished title = ask again.
//
// CHANGE LOG
// - 2026-10-13 : Initial version, ported 1:1 from AssignmentItem.cs and
//       SheetSetup.cs.
//     * isOverdue / statusText take "today" as a PARAMETER ("YYYY-MM-DD").
//       WHY: C# used the PC clock (DateTime.Today, India time). The nightly
//       job runs on a UTC server, so "today" is computed once, in India
//       time, by the caller and passed in - never read from the server clock
//       inside the logic.
//     * key keeps the C# format exactly (role as its enum NAME, date as
//       yyyyMMdd), so keys are comparable with Excel_Notify's sent logs.
// ================================================================

import { roleKey } from './roles.js';

export class AssignmentItem {
  constructor(fields) {
    this.sheetName = fields.sheetName;
    /** Discipline section, e.g. "MECHANICAL"; empty if none above the row. */
    this.discipline = fields.discipline || '';
    /** Sub-group heading, e.g. "AIR CONDITIONING SYSTEM LAYOUT"; may be empty. */
    this.subGroup = fields.subGroup || '';
    this.drawingNumber = fields.drawingNumber || '';
    this.drawingName = fields.drawingName || '';
    /** Role value (roles.js). */
    this.role = fields.role;
    /** Stage title the role was read from, e.g. "Modeling with Coordination". */
    this.stageTitle = fields.stageTitle;
    /** One person's name as typed (trimmed), e.g. "Abhi". */
    this.personName = fields.personName;
    /** Text in brackets after the name, e.g. "markup"; may be empty. */
    this.note = fields.note || '';
    /** Due date "YYYY-MM-DD"; null when the cell is empty or not a date. */
    this.dueDate = fields.dueDate || null;
    this.isDone = Boolean(fields.isDone);
    /** REMARKS column text; may be empty. */
    this.remarks = fields.remarks || '';
    /** Address of the Name cell, e.g. "F35" (for warnings and checks). */
    this.cellAddress = fields.cellAddress;
  }

  /** Not done AND due before today ("YYYY-MM-DD", India time - see change log). */
  isOverdue(today) {
    return !this.isDone && this.dueDate !== null && this.dueDate < today; // ISO strings sort as dates
  }

  /** "Done" / "Overdue" / "Pending" - text for emails and reports. */
  statusText(today) {
    if (this.isDone) return 'Done';
    return this.isOverdue(today) ? 'Overdue' : 'Pending';
  }

  /** Sent-log key without the email - same format as C# AssignmentItem.Key. */
  get key() {
    const drawing = this.drawingNumber ? this.drawingNumber : `${this.drawingName}@${this.cellAddress}`;
    const date = this.dueDate ? this.dueDate.replace(/-/g, '') : 'nodate';
    return `${this.sheetName}|${drawing}|${roleKey(this.role)}|${date}`;
  }
}

/** The lead's choice for ONE sheet (mirrors SheetSetup.cs). */
export class SheetSetup {
  constructor(sheetName) {
    this.sheetName = sheetName;
    /** Ticked roles (Set of Role values). */
    this.roles = new Set();
    /** Role -> stage block TITLE used for it. */
    this.blockTitles = new Map();
  }
}
