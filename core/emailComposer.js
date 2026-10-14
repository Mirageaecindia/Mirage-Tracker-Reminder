// ================================================================
// emailComposer.js
// ----------------------------------------------------------------
// PURPOSE : Build tonight's reminder email for ONE person x ONE sheet.
//           Only builds text (SRP) - sends nothing.
//           MIRRORS Excel_Notify\Services\EmailComposer.cs + EmailBranding.cs
//           (same tables, colours, grouping and logo).
//
// LAYOUT  : Subject : Reminder: drawings due today - <sheet> - 3 due today (1 overdue) - <workbook>
//           Hi <name>,
//           [sent on behalf of <person>]            (On-behalf mode only)
//           [summary] Sheet | Workbook | Due today | Overdue | Date | Questions to
//           DUE TODAY  -> per ROLE -> per DISCIPLINE table
//           OVERDUE    -> per ROLE -> per DISCIPLINE table (red rows)
//               # | Drawing No. | Drawing Name | Sub-group | Due | Status | Note | Remarks
//           footer + Mirage AEC logo
//
// SENDER  : always FROM project-reminders@ (the only mailbox the job may use).
//           Default   : Reply-To = the tracker's lead
//           On behalf : Reply-To = the person who chose it on the admin page,
//                       plus a "Sent on behalf of" line (user decision
//                       2026-10-13; a custom From display name is not relied on).
//
// CHANGE LOG
// - 2026-10-14 : Initial version, adapted from EmailComposer.cs.
//     * Kept: grouping Role -> Discipline, role heading shows the sheet's own
//       column title ("Modelers - Modeling with Coordination"), empty roles /
//       disciplines left out, overdue rows red, "No due date" text.
//     * NEW: two sections (Due today / Overdue) - a reminder must say clearly
//       which items are due TONIGHT and which were already late.
//     * Logo: same file and size as EmailBranding.cs (mirage_logo.jpg, 260 px
//       wide, inline cid:mirage_logo). The job attaches the bytes; this file
//       only writes the <img> and says which attachment is needed.
//     * Last line of defence: every To / Reply-To must be @mirageaec.com,
//       otherwise compose throws (the sender checks again).
// ================================================================

import * as H from './htmlBuilder.js';
import { roleShortName } from './roles.js';
import { ReminderType } from './sentLog.js';
import { isAllowedAddress } from './recipientResolver.js';

export const SenderMode = Object.freeze({
  Default: 'default',
  OnBehalf: 'onBehalf',
});

export const LOGO = Object.freeze({ contentId: 'mirage_logo', fileName: 'mirage_logo.jpg', contentType: 'image/jpeg', widthPx: 260 });

const HEADERS = ['#', 'Drawing No.', 'Drawing Name', 'Sub-group', 'Due', 'Status', 'Note', 'Remarks'];

/**
 * @param {object} recipient  one row from recipientBuilder (sheetName, email, displayName, names, entries)
 * @param {object} ctx
 * @param {string} ctx.workbookName
 * @param {string} ctx.today                     "YYYY-MM-DD"
 * @param {string} ctx.senderMode                SenderMode
 * @param {{name: string, email: string}} ctx.lead        tracker lead (Default Reply-To)
 * @param {{name: string, email: string}} [ctx.onBehalfOf] person for On-behalf mode
 * @param {boolean} [ctx.includeLogo]
 * @returns {{to: object[], replyTo: object[], subject: string, html: string, inlineImages: object[]}}
 */
export function composeReminder(recipient, ctx) {
  const due = recipient.entries.filter((e) => e.type === ReminderType.Due).map((e) => e.item);
  const overdue = recipient.entries.filter((e) => e.type === ReminderType.Overdue).map((e) => e.item);
  const contact = ctx.senderMode === SenderMode.OnBehalf && ctx.onBehalfOf ? ctx.onBehalfOf : ctx.lead;
  const name = recipient.displayName || recipient.names[0];

  const mail = {
    to: [{ name, email: recipient.email }],
    replyTo: contact ? [{ name: contact.name, email: contact.email }] : [],
    subject: subjectFor(recipient.sheetName, due.length, overdue.length, ctx.workbookName),
    html: '',
    inlineImages: ctx.includeLogo ? [{ ...LOGO }] : [],
  };
  ensureCompanyAddresses(mail);

  const html = [];
  H.beginBody(html);
  H.paragraph(html, `Hi ${name},`);
  if (ctx.senderMode === SenderMode.OnBehalf && ctx.onBehalfOf) {
    H.paragraph(html, `Sent on behalf of ${ctx.onBehalfOf.name} (${ctx.onBehalfOf.email}).`);
  }
  H.paragraph(html, due.length > 0
    ? 'This is your automatic reminder: the drawings below in the project tracker are due today.'
    : 'This is your automatic reminder: the drawings below in the project tracker are past their due date.');

  H.labelValueTable(html, [
    ['Sheet', recipient.sheetName],
    ['Workbook', ctx.workbookName],
    ['Due today', String(due.length)],
    ['Overdue', String(overdue.length)],
    ['Date', H.date(ctx.today, '-')],
    ['Questions to', contact ? `${contact.name} <${contact.email}>` : '-'],
  ]);

  if (due.length > 0) appendSection(html, 'Due today', due, ctx.today);
  if (overdue.length > 0) {
    appendSection(html, 'Overdue', overdue, ctx.today);
    H.paragraph(html, 'Overdue drawings are reminded once. If a drawing is finished, tick its Status in the tracker.');
  }

  H.endBody(html, `Automatic reminder from Mirage AEC Tracker Reminder. Reply to this email to reach ${contact ? contact.name : 'the tracker lead'}.`);
  if (ctx.includeLogo) {
    html.push(`<p><img src="cid:${LOGO.contentId}" width="${LOGO.widthPx}" alt="Mirage AEC" style="display:block;border:0;"></p>`);
  }
  mail.html = html.join('');
  return mail;
}

/** "Reminder: drawings due today - EW08 (100%) - 3 due today (1 overdue) - Project Tracker - MEP.xlsx" */
function subjectFor(sheet, dueCount, overdueCount, workbook) {
  if (dueCount > 0) {
    return `Reminder: drawings due today - ${sheet} - ${dueCount} due today`
      + (overdueCount > 0 ? ` (${overdueCount} overdue)` : '') + ` - ${workbook}`;
  }
  return `Reminder: overdue drawings - ${sheet} - ${overdueCount} overdue - ${workbook}`;
}

/** Section heading, then Role -> Discipline tables (as C# BuildBody). */
function appendSection(html, title, items, today) {
  H.heading(html, `${title} (${items.length})`);
  for (const [role, roleItems] of groupBy(items, (i) => i.role)) {
    const columnTitles = [...new Set(roleItems.map((i) => i.stageTitle))].join(', ');
    H.subHeading(html, `${roleShortName(role)} - ${columnTitles} (${roleItems.length})`);
    for (const [discipline, list] of groupBy(roleItems, (i) => i.discipline || '')) {
      if (discipline) H.paragraph(html, discipline);
      appendTable(html, list, today);
    }
  }
}

function appendTable(html, items, today) {
  const rows = items.map((item, index) => [
    String(index + 1), item.drawingNumber, item.drawingName, item.subGroup,
    H.date(item.dueDate, 'No due date'), item.statusText(today), item.note, item.remarks,
  ]);
  H.table(html, HEADERS, rows, (index) => items[index].isOverdue(today));
}

/** Throws if any To / Reply-To is outside the company (see change log). */
function ensureCompanyAddresses(mail) {
  for (const r of [...mail.to, ...mail.replyTo]) {
    if (!isAllowedAddress(r.email)) throw new Error(`Blocked: "${r.email}" is not a company address.`);
  }
}

/** Map preserving first-seen order (items are already sorted role -> discipline). */
function groupBy(list, keyOf) {
  const map = new Map();
  for (const x of list) {
    const key = keyOf(x);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(x);
  }
  return map;
}
