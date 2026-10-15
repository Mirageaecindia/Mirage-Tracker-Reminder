// ================================================================
// recipientBuilder.js
// ----------------------------------------------------------------
// PURPOSE : Turn the assignments read from ONE tracker into tonight's
//           reminder emails: one email = one person x one sheet (same as
//           Excel_Notify). Pure in-memory logic: no Graph, no SharePoint (SRP).
//           MIRRORS Excel_Notify\Services\RecipientBuilder.cs - grouping,
//           merging and sorting are the same; WHICH items are included is
//           the nightly rule below instead of the lead's window filters.
//
// NIGHTLY RULE (user decisions 2026-10-13 / 14), per item:
//   done (Status ticked / COMPLETED)          -> never
//   no due date (empty or bad date text)      -> never (bad text is already
//                                                a warning in the lead summary)
//   due date = today                          -> DUE reminder
//   due date before today, not done           -> OVERDUE reminder, ONCE
//   due date after today                      -> nothing yet
//   already in the sent log for that address  -> skipped (counted)
//
// BACKLOG (overdue items due BEFORE the tracker was registered) - per-tracker
//   choice on the admin page:
//   afterRegistration : backlog never emailed (default)
//   emailOnce         : backlog emailed once like any overdue item
//   leadSummaryOnly   : backlog listed in the lead summary, not emailed
//
// CHANGE LOG
// - 2026-10-14 : Initial version, adapted from RecipientBuilder.cs.
//     * Kept: group by sheet + tracker name (case/space-insensitive); names
//       matched to the SAME address in the same sheet are MERGED into one
//       email ("Ruwaisid" + "Ruwaishid"); unmatched names are never emailed;
//       items sorted role -> discipline -> drawing number.
//     * Replaced: the window's Show / "only not emailed before" filters by
//       the nightly rule above; unmatched rows go to the lead summary.
//     * Roles: only the roles the lead ticked for each sheet are read in the
//       first place (trackerReader + SheetSetup), so no role filter here.
// - 2026-10-15 : ADDED resendPolicy (user request: per-tracker radio on the admin page).
//     WHY : until now an item was emailed to a person ONCE (sent-log check) - a lead
//           who wants reminders to keep coming until the drawing is done had no way.
//     WHAT: new ResendPolicy + parameter resendPolicy of buildReminders():
//           once     : old behaviour - skip if already in the sent log (default, so
//                      trackers saved before this change behave exactly as before)
//           daily    : resend every day, but skip what was already sent TODAY
//                      (a manual same-day re-run never duplicates) - uses
//                      sentLog.wasSentOn() (added in sentLog.js)
//           everyRun : resend on every run, no check at all (a manual re-run sends again)
//     NOTE: the alreadySentCount in the summary only counts skipped items, so it is
//           0 in everyRun mode by design.
// ================================================================

import { ReminderType } from './sentLog.js';

export const BacklogPolicy = Object.freeze({
  AfterRegistration: 'afterRegistration',
  EmailOnce: 'emailOnce',
  LeadSummaryOnly: 'leadSummaryOnly',
});

/** Per-tracker choice: how often the same item may be emailed to the same person (see change log 2026-10-15). */
export const ResendPolicy = Object.freeze({
  Once: 'once',
  Daily: 'daily',
  EveryRun: 'everyRun',
});

/**
 * @param {object} args
 * @param {import('./assignmentItem.js').AssignmentItem[]} args.items
 * @param {string} args.today          "YYYY-MM-DD", India time
 * @param {string} args.trackerId
 * @param {string} args.registeredOn   "YYYY-MM-DD" the tracker was registered
 * @param {string} [args.backlogPolicy]
 * @param {string} [args.resendPolicy]  ResendPolicy; missing/unknown -> once (old behaviour)
 * @param {import('./recipientResolver.js').RecipientResolver} args.resolver
 * @param {import('./sentLog.js').SentLog} args.sentLog
 */
export function buildReminders({ items, today, trackerId, registeredOn, backlogPolicy, resendPolicy, resolver, sentLog }) {
  const policy = Object.values(BacklogPolicy).includes(backlogPolicy) ? backlogPolicy : BacklogPolicy.AfterRegistration;
  const resend = Object.values(ResendPolicy).includes(resendPolicy) ? resendPolicy : ResendPolicy.Once;
  const result = {
    /** emails to send */
    recipients: [],
    /** names with reminders but no sure email match */
    unmatched: [],
    /** backlog items for the lead summary (leadSummaryOnly) */
    backlog: [],
    /** items skipped because already emailed */
    alreadySentCount: 0,
    /** backlog items not emailed (afterRegistration) */
    backlogSkippedCount: 0,
  };

  // ---- 1. pick tonight's items + their type ----
  const candidates = [];
  for (const item of items) {
    const type = reminderTypeFor(item, today);
    if (!type) continue;

    const isBacklog = type === ReminderType.Overdue && item.dueDate < registeredOn;
    if (isBacklog && policy === BacklogPolicy.AfterRegistration) {
      result.backlogSkippedCount++;
      continue;
    }
    if (isBacklog && policy === BacklogPolicy.LeadSummaryOnly) {
      result.backlog.push(item);
      continue;
    }
    candidates.push({ item, type });
  }

  // ---- 2. group by sheet + name, resolve, merge by address ----
  const rows = new Map();
  const groups = groupBy(candidates, (c) => `${c.item.sheetName}|${nameKey(c.item.personName)}`);

  for (const group of groups.values()) {
    const first = group[0].item;
    const user = resolver.resolve(first.personName);

    if (!user) {
      result.unmatched.push({
        sheetName: first.sheetName,
        name: first.personName.trim(),
        dueCount: group.filter((c) => c.type === ReminderType.Due).length,
        overdueCount: group.filter((c) => c.type === ReminderType.Overdue).length,
      });
      continue; // never email a guess
    }

    const rowKey = `${first.sheetName}|${user.email.toLowerCase()}`;
    if (!rows.has(rowKey)) {
      rows.set(rowKey, { sheetName: first.sheetName, email: user.email, displayName: user.displayName, names: [], entries: [] });
    }
    const row = rows.get(rowKey);

    const spelling = first.personName.trim();
    if (!row.names.some((n) => n.toLowerCase() === spelling.toLowerCase())) row.names.push(spelling);

    for (const c of group) {
      if (alreadySent(sentLog, resend, trackerId, c, user.email, today)) result.alreadySentCount++;
      else row.entries.push(c);
    }
  }

  // ---- 3. keep rows with something to send, sort like C# ----
  for (const row of rows.values()) {
    if (row.entries.length === 0) continue;
    row.entries.sort(compareEntries);
    result.recipients.push(row);
  }
  result.recipients.sort((a, b) => a.sheetName.localeCompare(b.sheetName) || a.names.join(', ').localeCompare(b.names.join(', ')));
  result.unmatched.sort((a, b) => a.sheetName.localeCompare(b.sheetName) || a.name.localeCompare(b.name));
  return result;
}

/** True = skip this item for this person, according to the tracker's resend policy (see change log 2026-10-15). */
function alreadySent(sentLog, resend, trackerId, entry, email, today) {
  if (resend === ResendPolicy.EveryRun) return false;
  if (resend === ResendPolicy.Daily) return sentLog.wasSentOn(trackerId, entry.item.key, entry.type, email, today);
  return sentLog.wasSent(trackerId, entry.item.key, entry.type, email);
}

/** "due" / "overdue" / null for one item (see NIGHTLY RULE). */
export function reminderTypeFor(item, today) {
  if (item.isDone || !item.dueDate) return null;
  if (item.dueDate === today) return ReminderType.Due;
  if (item.dueDate < today) return ReminderType.Overdue;
  return null;
}

/** Role, then discipline, then drawing number (stable email tables, as C#). */
function compareEntries(a, b) {
  return a.item.role - b.item.role
    || a.item.discipline.localeCompare(b.item.discipline, undefined, { sensitivity: 'base' })
    || a.item.drawingNumber.localeCompare(b.item.drawingNumber, undefined, { sensitivity: 'base' });
}

/** "  AKASH  kumar " -> "akash kumar" (grouping key only, as C# NameKey). */
function nameKey(name) {
  return String(name ?? '').trim().split(/\s+/).join(' ').toLowerCase();
}

function groupBy(list, keyOf) {
  const map = new Map();
  for (const x of list) {
    const key = keyOf(x);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(x);
  }
  return map;
}
