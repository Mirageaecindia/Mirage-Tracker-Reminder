// ================================================================
// sentLog.js
// ----------------------------------------------------------------
// PURPOSE : Remember which reminder was emailed to which person, so a
//           re-run (manual "Run now", GitHub retry) never sends twice and an
//           overdue item gets its overdue email only ONCE.
//           Pure data + rules (SRP). Loading / saving the JSON file in
//           SharePoint is the job's I/O code.
//           MIRRORS Excel_Notify\Services\SentLogStore.cs (rules adapted to
//           a nightly job - see change log).
//
// FILE    : TrackerReminder\sentlog.json in the Automation site:
//           { "version": 1, "entries": [ { "trackerId", "itemKey", "type",
//             "email", "sentAt" }, ... ] }
//
// CHANGE LOG
// - 2026-10-14 : Initial version, adapted from SentLogStore.cs.
//     * Record key = trackerId + AssignmentItem.key (sheet + drawing + role
//       + due date, same as C#) + TYPE + email. NEW: type ("due" /
//       "overdue") - the same item is legitimately emailed twice: once on
//       its due day, once (only once) when it becomes overdue.
//     * A changed due date or a different person = a different key = a new
//       email (same user decision as Excel_Notify).
//     * ONE file for all trackers (trackerId in every entry) instead of one
//       file per workbook: the job reads/writes one SharePoint file per run.
//     * Pruning (C# deleted logs unused for 30 days; a nightly job is never
//       "unused", so entries are pruned by age instead):
//         - "due" entries after 7 days. WHY: a due reminder only ever goes
//           out ON the due date, so an old one can never be re-sent anyway;
//           the record only stops same-day re-runs.
//         - "overdue" entries after 365 days. WHY: they are what stops a
//           second overdue email; dropping them early would re-send. After a
//           year the item is either done or long abandoned (known limit).
// - 2026-10-15 : ADDED wasSentOn() and "re-record refreshes sentAt" (user request:
//       per-tracker resend option, see recipientBuilder.js ResendPolicy).
//     WHY : resend policy "daily" must know whether an item already went out
//           TODAY (India date), so a manual same-day re-run does not duplicate.
//     WHAT: * index is now a Map (key -> entry) instead of a Set, to find the
//             entry's sentAt without scanning the whole list.
//           * wasSentOn(..., istDate): true if the entry exists AND its sentAt
//             falls on that India date (sentAt is UTC; +05:30 applied here).
//           * record(): an existing key now gets its sentAt refreshed instead of
//             being ignored. Old behaviour is unchanged for "once" trackers
//             (they never re-record a key they skipped).
//     NOTE: refreshing sentAt also restarts the prune clock - a resent item stays
//           in the log while it keeps being sent, which is what we want.
// ================================================================

export const ReminderType = Object.freeze({
  Due: 'due',
  Overdue: 'overdue',
});

const KEEP_DAYS = Object.freeze({ due: 7, overdue: 365 });

export class SentLog {
  /** @param {{version?: number, entries?: object[]}} [data]  parsed sentlog.json (missing/corrupt -> empty) */
  constructor(data) {
    this.entries = Array.isArray(data && data.entries) ? data.entries.filter(isValidEntry) : [];
    this.index = new Map(this.entries.map((e) => [indexKey(e.trackerId, e.itemKey, e.type, e.email), e]));
  }

  /** True if this reminder already went to this address. */
  wasSent(trackerId, itemKey, type, email) {
    return this.index.has(indexKey(trackerId, itemKey, type, email));
  }

  /** True if this reminder went to this address on the given India date "YYYY-MM-DD" (see change log 2026-10-15). */
  wasSentOn(trackerId, itemKey, type, email, istDate) {
    const entry = this.index.get(indexKey(trackerId, itemKey, type, email));
    return Boolean(entry) && istDateOf(entry.sentAt) === istDate;
  }

  /** Records one email's items (call only AFTER the email was accepted by Graph). An existing record gets its sentAt refreshed. */
  record(trackerId, itemKeys, type, email, sentAtIso) {
    for (const itemKey of itemKeys) {
      const key = indexKey(trackerId, itemKey, type, email);
      const existing = this.index.get(key);
      if (existing) {
        existing.sentAt = sentAtIso;
        continue;
      }
      const entry = { trackerId, itemKey, type, email: email.trim(), sentAt: sentAtIso };
      this.index.set(key, entry);
      this.entries.push(entry);
    }
  }

  /** Drops old entries (see change log). Returns how many were removed. */
  prune(nowIso) {
    const now = Date.parse(nowIso);
    const before = this.entries.length;
    this.entries = this.entries.filter((e) => {
      const keepDays = KEEP_DAYS[e.type] ?? KEEP_DAYS.overdue;
      return now - Date.parse(e.sentAt) <= keepDays * 86400000;
    });
    this.index = new Map(this.entries.map((e) => [indexKey(e.trackerId, e.itemKey, e.type, e.email), e]));
    return before - this.entries.length;
  }

  /** Data for sentlog.json. */
  toJSON() {
    return { version: 1, entries: this.entries };
  }
}

const IST_OFFSET_MS = 330 * 60000; // India = UTC+05:30

/** UTC ISO timestamp -> India date "YYYY-MM-DD". */
function istDateOf(iso) {
  return new Date(Date.parse(iso) + IST_OFFSET_MS).toISOString().slice(0, 10);
}

function indexKey(trackerId, itemKey, type, email) {
  return `${trackerId}|${itemKey}|${type}|${String(email).trim().toLowerCase()}`;
}

function isValidEntry(e) {
  return e && e.trackerId && e.itemKey && e.type && e.email && !Number.isNaN(Date.parse(e.sentAt));
}
