// ================================================================
// leadSummaryComposer.js
// ----------------------------------------------------------------
// PURPOSE : Build the nightly digest for ONE tracker's lead: what was sent,
//           what failed, which names could not be matched to an email, the
//           backlog (if the tracker uses "lead summary only"), bad cells and
//           job errors. Only builds text (SRP).
//           NEW - closest Excel_Notify equivalent is ReportComposer.cs
//           (same overview-table-then-details layout).
//
// CHANGE LOG
// - 2026-10-14 : Initial version (user decision 2026-10-13: nightly summary
//       to the lead instead of CC on every reminder).
//     * Unmatched names come with the action to take ("map it on the admin
//       page"), because nobody is at a window at 8 PM to click Pick...
//     * Sent even when NO reminder went out but there are unmatched names,
//       warnings or errors - silence must mean "all fine" (decided by the
//       caller via shouldSend()).
//     * TEST MODE banner at the top when the job is in test mode, so a test
//       summary can never be mistaken for a real one.
// ================================================================

import * as H from './htmlBuilder.js';
import { roleShortName } from './roles.js';
import { ReminderType } from './sentLog.js';
import { isAllowedAddress } from './recipientResolver.js';
import { LOGO } from './emailComposer.js';

/**
 * @param {object} s
 * @param {string} s.trackerName
 * @param {string} s.workbookName
 * @param {string} s.today                         "YYYY-MM-DD"
 * @param {{name: string, email: string}} s.lead
 * @param {{recipient: object, ok: boolean, error?: string}[]} s.results   one per reminder email
 * @param {object[]} s.unmatched                   from recipientBuilder
 * @param {object[]} s.backlog                     AssignmentItems (leadSummaryOnly)
 * @param {number} s.alreadySentCount
 * @param {{sheetName: string, cell: string, message: string}[]} s.warnings
 * @param {string[]} s.errors                      job-level problems (file not readable, ...)
 * @param {boolean} [s.testMode]
 * @param {boolean} [s.includeLogo]
 */
export function composeLeadSummary(s) {
  if (!isAllowedAddress(s.lead.email)) throw new Error(`Blocked: lead "${s.lead.email}" is not a company address.`);

  const sent = s.results.filter((r) => r.ok);
  const failed = s.results.filter((r) => !r.ok);
  const count = (list, type) => list.reduce((n, r) => n + r.recipient.entries.filter((e) => e.type === type).length, 0);

  const html = [];
  H.beginBody(html);
  if (s.testMode) {
    html.push('<p style="background:#fff4ce;border:1px solid #c8a400;padding:6px 10px;font-weight:bold;">'
      + 'TEST MODE - reminders below were sent to the test address only, not to the people listed.</p>');
  }
  H.paragraph(html, `Hi ${s.lead.name},`);
  H.paragraph(html, `Tonight's reminder run for ${s.trackerName}.`);

  H.labelValueTable(html, [
    ['Tracker', s.trackerName],
    ['Workbook', s.workbookName],
    ['Date', H.date(s.today, '-')],
    ['Emails sent', String(sent.length)],
    ['Emails failed', String(failed.length)],
    ['Due-today items reminded', String(count(sent, ReminderType.Due))],
    ['Overdue items reminded (once)', String(count(sent, ReminderType.Overdue))],
    ['Already reminded before (skipped)', String(s.alreadySentCount)],
    ['Names not matched', String(s.unmatched.length)],
    ['Cell warnings', String(s.warnings.length)],
  ]);

  if (s.errors.length > 0) {
    H.heading(html, `Problems (${s.errors.length})`);
    s.errors.forEach((e) => H.paragraph(html, e));
  }

  if (s.unmatched.length > 0) {
    H.heading(html, `Names not matched to an email (${s.unmatched.length}) - NOT reminded`);
    H.paragraph(html, 'Map these names to a person on the Tracker Reminder admin page; they will be reminded from the next run.');
    H.table(html, ['Sheet', 'Name in tracker', 'Due today', 'Overdue'],
      s.unmatched.map((u) => [u.sheetName, u.name, String(u.dueCount), String(u.overdueCount)]));
  }

  if (s.results.length > 0) {
    H.heading(html, `Reminder emails (${s.results.length})`);
    H.table(html, ['Name(s) in tracker', 'Email', 'Sheet', 'Due today', 'Overdue', 'Result'],
      s.results.map((r) => [
        r.recipient.names.join(', '), r.recipient.email, r.recipient.sheetName,
        String(r.recipient.entries.filter((e) => e.type === ReminderType.Due).length),
        String(r.recipient.entries.filter((e) => e.type === ReminderType.Overdue).length),
        r.ok ? 'Sent' : `FAILED: ${r.error || 'unknown error'}`,
      ]),
      (index) => !s.results[index].ok);
  }

  if (s.backlog.length > 0) {
    H.heading(html, `Overdue from before registration (${s.backlog.length}) - not emailed`);
    H.table(html, ['Sheet', 'Person', 'Role', 'Drawing No.', 'Drawing Name', 'Due'],
      s.backlog.map((i) => [i.sheetName, i.personName, roleShortName(i.role), i.drawingNumber, i.drawingName, H.date(i.dueDate, '-')]));
  }

  if (s.warnings.length > 0) {
    H.heading(html, `Cell warnings (${s.warnings.length})`);
    H.table(html, ['Sheet', 'Cell', 'Problem'], s.warnings.map((w) => [w.sheetName, w.cell, w.message]));
  }

  H.endBody(html, 'Automatic summary from Mirage AEC Tracker Reminder.');
  if (s.includeLogo) {
    html.push(`<p><img src="cid:${LOGO.contentId}" width="${LOGO.widthPx}" alt="Mirage AEC" style="display:block;border:0;"></p>`);
  }

  return {
    to: [{ name: s.lead.name, email: s.lead.email }],
    replyTo: [],
    subject: `${s.testMode ? '[TEST] ' : ''}Tracker reminders - ${s.trackerName} - ${H.date(s.today, '')} - `
      + `${sent.length} sent${failed.length ? `, ${failed.length} failed` : ''}${s.unmatched.length ? `, ${s.unmatched.length} not matched` : ''}`,
    html: html.join(''),
    inlineImages: s.includeLogo ? [{ ...LOGO }] : [],
  };
}

/** Send the summary only when there is something to report (see change log). */
export function shouldSendLeadSummary(s) {
  return s.results.length > 0 || s.unmatched.length > 0 || s.backlog.length > 0 || s.warnings.length > 0 || s.errors.length > 0;
}
