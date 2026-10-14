// ================================================================
// schedule.js
// ----------------------------------------------------------------
// PURPOSE : WHEN a tracker sends. Each tracker has its own send time in
//           India time (default 8:00 PM, 30-minute steps). GitHub wakes the
//           job every 30 minutes; this file answers "is this tracker due
//           now?" and remembers "already sent today" (runstate.json).
//           Pure logic + data (SRP); reading/saving the file is job code.
//
// RULE    : due = (India time now >= the tracker's send time)
//                 AND (the tracker has not run yet today, India date)
//           -> a skipped or late GitHub wake-up is caught up by the next one.
//
// FILE    : TrackerReminder\runstate.json in the Automation site:
//   { "version": 1,
//     "trackers": { "test-mep": { "lastRunOn": "2026-10-14", "lastRunAt": "<ISO>" } },
//     "lastAlert": { "on": "2026-10-14", "key": "<problem text>" } }
//
// CHANGE LOG
// - 2026-10-14 : Initial version (user decision 2026-10-14: per-tracker
//       send time, editable in IST, default 8 PM, 30-minute steps).
//     * Only HH:00 and HH:30 allowed. WHY: the job wakes at :00 and :30
//       India time; any other minute would silently mean "the next half hour".
//     * India time is computed from UTC + 5:30 (India has no daylight
//       saving), never from the server's clock zone (GitHub runs in UTC).
//     * Known limit: if GitHub misses EVERY wake-up between the send time
//       and midnight, that day's due reminders are not sent; the next day
//       they are overdue and get their one overdue email instead.
//     * lastAlert: the same problem text alerts at most ONCE per India day
//       (otherwise a broken file would email the admin 48 times a day).
// ================================================================

export const DEFAULT_SEND_TIME = '20:00';
export const SEND_TIME = /^([01]\d|2[0-3]):(00|30)$/;

const IST_OFFSET_MS = 5.5 * 3600 * 1000;

/** India date + minutes since midnight, from a UTC timestamp. */
export function indiaNow(nowMs = Date.now()) {
  const ist = new Date(nowMs + IST_OFFSET_MS);
  return {
    date: ist.toISOString().slice(0, 10),            // "2026-10-14"
    minutes: ist.getUTCHours() * 60 + ist.getUTCMinutes(),
    label: ist.toISOString().slice(11, 16),           // "20:00"
  };
}

/** "20:30" -> 1230 minutes. */
export function minutesOf(sendTime) {
  const [h, m] = sendTime.split(':').map(Number);
  return h * 60 + m;
}

/** "20:00" -> "8:00 PM" (for screens and logs). */
export function sendTimeLabel(sendTime) {
  const [h, m] = sendTime.split(':').map(Number);
  return `${((h + 11) % 12) + 1}:${String(m).padStart(2, '0')} ${h < 12 ? 'AM' : 'PM'}`;
}

/** See RULE. */
export function isTrackerDue(sendTime, now, lastRunOn) {
  return lastRunOn !== now.date && now.minutes >= minutesOf(sendTime || DEFAULT_SEND_TIME);
}

/** runstate.json contents (see FILE). Missing / corrupt parts -> empty. */
export class RunState {
  constructor(data) {
    this.trackers = data && typeof data.trackers === 'object' && data.trackers ? { ...data.trackers } : {};
    this.lastAlert = data && data.lastAlert ? { ...data.lastAlert } : null;
    this.changed = false;
  }

  lastRunOn(trackerId) {
    return this.trackers[trackerId] ? this.trackers[trackerId].lastRunOn : null;
  }

  markRun(trackerId, date, atIso) {
    this.trackers[trackerId] = { lastRunOn: date, lastRunAt: atIso };
    this.changed = true;
  }

  /** True if this exact problem text was already alerted today. */
  alreadyAlerted(date, key) {
    return Boolean(this.lastAlert && this.lastAlert.on === date && this.lastAlert.key === key);
  }

  markAlerted(date, key) {
    this.lastAlert = { on: date, key };
    this.changed = true;
  }

  toJSON() {
    return { version: 1, trackers: this.trackers, lastAlert: this.lastAlert };
  }
}
