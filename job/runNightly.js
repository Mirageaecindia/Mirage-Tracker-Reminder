// ================================================================
// job/runNightly.js
// ----------------------------------------------------------------
// PURPOSE : One wake-up of the reminder job (every 30 min). ONLY orchestrates (like Excel_Notify's
//           NotifyCommand): every decision lives in core\, every Graph call
//           in job\*.js.
//
// FLOW    : 1. settings (job/config.js), app sign-in
//           2. trackers.json + runstate.json -> which trackers are DUE now
//              (core/schedule.js). None due and no problem -> stop here.
//           3. name_map.json, sentlog.json, company directory
//           4. for EACH due tracker (one failing never stops the others):
//                download -> snapshot -> layouts -> read (lead's setup)
//                -> People sheet -> resolve names -> decide (due / overdue)
//                -> send each reminder -> record in sent log
//                -> lead summary (only when there is something to report)
//                then mark it "ran today" in runstate.json
//           5. save sentlog.json + runstate.json once, at the end
//           6. anything went wrong -> alert email to ALERT_EMAIL (once per
//              day per problem) and exit code 1 (GitHub marks the run red)
//
// RUN     : GitHub Actions (.github/workflows/nightly.yml), or locally:
//             node --env-file=.env job/runNightly.js
//           DRY_RUN=true -> decides and prints, sends nothing, writes nothing.
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * "today" = India date (UTC+05:30). GitHub's server clock is UTC; at
//       8 PM IST it is 14:30 UTC on the SAME day, but computing India time
//       explicitly keeps it right even if the schedule is ever moved past
//       midnight UTC.
//     * TEST MODE = GLOBAL_TEST_MODE (default on) OR the tracker's own
//       testMode. Test sends are logged under "<id>#test", so switching a
//       tracker live does not skip reminders that were only test-mailed.
//     * The sent log is recorded per email right after Graph accepts it and
//       saved ONCE at the end. Known limit: if the run crashes after sending
//       but before saving, the next run may re-send those (better than the
//       opposite - losing reminders).
//     * A hidden People sheet is reported (the snapshot skips hidden sheets).
//     * The lead summary also goes to the test address in test mode.
// - 2026-10-14 : Per-tracker SEND TIME (user decision: IST, default 8 PM,
//       30-minute steps). The workflow now wakes every 30 minutes.
//     * Only DUE trackers run (core/schedule.js). When nothing is due the
//       run stops before loading the directory / sent log - a few seconds.
//     * A tracker is marked "ran today" after its run EVEN IF it failed.
//       WHY: otherwise a moved file would be retried - and alerted - every
//       30 minutes until midnight. The alert says how to retry by hand.
//     * Manual runs with IGNORE_SEND_TIME run every tracker and do NOT mark
//       "ran today". WHY: a noon test must not cancel that day's real run
//       (the sent log still stops duplicate reminders).
//     * Alerts: the same problem text emails at most ONCE per India day; a
//       repeat is only logged and the run stays green (GitHub would also
//       email the owner about every red run - 48 times a day).
//     * "today" now comes from schedule.indiaNow() (one place for IST maths);
//       the local indiaToday() helper was removed.
//     * BUG FOUND IN TEST (same day): runstate.json was read AFTER
//       trackers.json. With a broken trackers.json the read threw first, so
//       the run continued with an EMPTY run state: the alert was not
//       throttled, and saving the alert mark OVERWROTE runstate.json, wiping
//       every tracker's "ran today" record. Fix: read runstate.json first.
// ================================================================

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join, posix } from 'node:path';
import XLSX from 'xlsx';

import { loadJobConfig } from './config.js';
import { GraphClient } from './graphClient.js';
import { SharePointFiles, CONFIG_FILES } from './sharePointFiles.js';
import { getCompanyDirectory } from './directoryService.js';
import { MailSender } from './mailSender.js';

import { takeSnapshot } from '../core/workbookSnapshot.js';
import { detectLayout } from '../core/sheetLayoutDetector.js';
import { readTracker } from '../core/trackerReader.js';
import { readPeopleSheet, findPeopleSheetName } from '../core/peopleSheetReader.js';
import { RecipientResolver, parseNameMap } from '../core/recipientResolver.js';
import { buildReminders } from '../core/recipientBuilder.js';
import { SentLog, ReminderType } from '../core/sentLog.js';
import { composeReminder, LOGO } from '../core/emailComposer.js';
import { composeLeadSummary, shouldSendLeadSummary } from '../core/leadSummaryComposer.js';
import { parseTrackersFile, toSheetSetups } from '../core/trackerConfig.js';
import * as H from '../core/htmlBuilder.js';
import { indiaNow, isTrackerDue, sendTimeLabel, RunState } from '../core/schedule.js';

const HERE = dirname(fileURLToPath(import.meta.url));

async function main() {
  const cfg = loadJobConfig(); // throws on missing settings -> caught below (no mail possible yet)
  const graph = new GraphClient(cfg);
  const files = new SharePointFiles(graph);
  const configPath = (name) => posix.join(cfg.configFolder, name);
  const read = (name) => files.readJson(cfg.configSiteHost, cfg.configSitePath, configPath(name));
  const write = (name, data) => files.writeJson(cfg.configSiteHost, cfg.configSitePath, configPath(name), data);
  const now = indiaNow();
  const today = now.date;
  const nowIso = new Date().toISOString();
  const problems = []; // everything that goes into the alert email
  const sender = new MailSender(graph, { senderMailbox: cfg.senderMailbox, logoBytes: readLogo(), dryRun: cfg.dryRun });
  let runState = new RunState(null);

  console.log(`Tracker Reminder wake-up - India ${today} ${now.label}`
    + `${cfg.dryRun ? ' - DRY RUN' : ''}${cfg.globalTestMode ? ' - GLOBAL TEST MODE' : ''}${cfg.ignoreSendTime ? ' - IGNORING SEND TIMES' : ''}`);

  try {
    // ---- 2. which trackers are due ----
    // runstate FIRST: if trackers.json is broken, the alert throttle and the
    // "ran today" records must still be known (bug found 2026-10-14 in test).
    runState = new RunState(await read(CONFIG_FILES.runState));
    const { trackers, problems: configProblems } = parseTrackersFile(await read(CONFIG_FILES.trackers));
    problems.push(...configProblems);

    const due = [];
    for (const t of trackers.filter((x) => x.enabled)) {
      if (cfg.ignoreSendTime || isTrackerDue(t.sendTime, now, runState.lastRunOn(t.id))) due.push(t);
      else console.log(`   ${t.name}: not due (sends at ${sendTimeLabel(t.sendTime)}, last ran ${runState.lastRunOn(t.id) || 'never'})`);
    }

    if (due.length > 0) {
      // ---- 3. shared data ----
      const { map: sharedMap, rejected } = parseNameMap((await read(CONFIG_FILES.nameMap)) || []);
      if (rejected) problems.push(`name_map.json: ${rejected} invalid entr${rejected === 1 ? 'y' : 'ies'} ignored (each needs Name + @mirageaec.com Email).`);
      const sentLog = new SentLog(await read(CONFIG_FILES.sentLog));
      const directory = await getCompanyDirectory(graph);
      console.log(`Due now: ${due.length} tracker(s). ${sharedMap.size} mapped name(s), ${directory.length} company user(s).`);

      // ---- 4. each due tracker ----
      for (const tracker of due) {
        try {
          await runTracker({ tracker, cfg, files, sender, sentLog, sharedMap, directory, today, nowIso });
        } catch (e) {
          problems.push(`${tracker.name}: ${e.message} (it runs again tomorrow; to retry today use Actions -> Run workflow)`);
          console.error(`!! ${tracker.name}: ${e.message}`);
        }
        if (!cfg.ignoreSendTime && !cfg.dryRun) runState.markRun(tracker.id, today, nowIso); // see change log
      }

      // ---- 5. save the sent log once ----
      const pruned = sentLog.prune(nowIso);
      if (!cfg.dryRun) await write(CONFIG_FILES.sentLog, sentLog.toJSON());
      console.log(`Sent log: ${sentLog.entries.length} record(s), ${pruned} old record(s) pruned${cfg.dryRun ? ' (not saved - dry run)' : ''}.`);
    } else if (problems.length === 0) {
      console.log('Nothing due now.');
    }
  } catch (e) {
    problems.push(`Run stopped: ${e.message}`);
    console.error(`!! Run stopped: ${e.message}`);
  }

  // ---- 6. alert (once per day per problem) ----
  if (problems.length) {
    const key = problems.join('\n');
    if (!cfg.ignoreSendTime && runState.alreadyAlerted(today, key)) {
      console.log('Same problem(s) already alerted today - not emailing again.');
    } else {
      await sendAlert(sender, cfg, today, problems).catch((e) => console.error(`!! Alert email failed too: ${e.message}`));
      if (!cfg.dryRun) runState.markAlerted(today, key);
      process.exitCode = 1;
    }
  }

  if (runState.changed && !cfg.dryRun) {
    await write(CONFIG_FILES.runState, runState.toJSON()).catch((e) => {
      console.error(`!! Could not save runstate.json: ${e.message}`);
      process.exitCode = 1;
    });
  }
}

/** One tracker, start to finish (see FLOW 3). */
async function runTracker({ tracker, cfg, files, sender, sentLog, sharedMap, directory, today, nowIso }) {
  const testMode = cfg.globalTestMode || tracker.testMode;
  const testRecipient = testMode ? (tracker.testRecipient || cfg.testRecipient) : null;
  const logId = testMode ? `${tracker.id}#test` : tracker.id;
  const workbookName = posix.basename(tracker.filePath);
  console.log(`\n== ${tracker.name}${testMode ? `  [TEST -> ${testRecipient}]` : ''}`);

  const bytes = await files.downloadFile(tracker.siteHost, tracker.sitePath, tracker.filePath);
  const snapshot = takeSnapshot(XLSX, bytes, { name: workbookName, fullName: `${tracker.sitePath}/${tracker.filePath}` });
  const layouts = new Map(snapshot.sheets.map((s) => [s.name, detectLayout(s)]));
  const read = readTracker(snapshot, layouts, toSheetSetups(tracker));
  const people = readPeopleSheet(snapshot);

  const errors = [];
  const hiddenPeople = findPeopleSheetName(snapshot.hiddenSheets);
  if (hiddenPeople) errors.push(`The "${hiddenPeople}" sheet is hidden, so its names were not used - unhide it.`);

  const resolver = new RecipientResolver(directory, [sharedMap, people.map]);
  const plan = buildReminders({
    items: read.items, today, trackerId: logId, registeredOn: tracker.registeredOn,
    backlogPolicy: tracker.backlogPolicy, resolver, sentLog,
  });

  const results = [];
  for (const recipient of plan.recipients) {
    try {
      const mail = composeReminder(recipient, {
        workbookName, today, senderMode: tracker.senderMode, lead: tracker.lead,
        onBehalfOf: tracker.onBehalfOf, includeLogo: true,
      });
      await sender.send(mail, testRecipient);
      results.push({ recipient, ok: true });
      if (!cfg.dryRun) {
        for (const type of [ReminderType.Due, ReminderType.Overdue]) {
          const keys = recipient.entries.filter((e) => e.type === type).map((e) => e.item.key);
          if (keys.length) sentLog.record(logId, keys, type, recipient.email, nowIso);
        }
      }
    } catch (e) {
      results.push({ recipient, ok: false, error: e.message });
    }
  }

  const summaryInput = {
    trackerName: tracker.name, workbookName, today, lead: tracker.lead, results,
    unmatched: plan.unmatched, backlog: plan.backlog, alreadySentCount: plan.alreadySentCount,
    warnings: [...read.warnings, ...people.warnings], errors, testMode, includeLogo: true,
  };
  if (shouldSendLeadSummary(summaryInput)) await sender.send(composeLeadSummary(summaryInput), testRecipient);

  const failed = results.filter((r) => !r.ok);
  console.log(`   items read ${read.items.length} | emails ${results.length - failed.length} sent, ${failed.length} failed`
    + ` | not matched ${plan.unmatched.length} | already sent ${plan.alreadySentCount} | warnings ${summaryInput.warnings.length}`);
  failed.forEach((f) => console.log(`   FAILED ${f.recipient.email}: ${f.error}`));
  if (failed.length) throw new Error(`${failed.length} reminder email(s) failed - see the lead summary.`);
}

/** "The job had problems" email to ALERT_EMAIL (never test-redirected: it IS for the admin). */
async function sendAlert(sender, cfg, today, problems) {
  const html = [];
  H.beginBody(html);
  H.paragraph(html, `The Tracker Reminder run on ${H.date(today, today)} had ${problems.length} problem(s):`);
  problems.forEach((p) => H.paragraph(html, `- ${p}`));
  H.paragraph(html, 'Check the run log in GitHub -> Actions -> "Nightly tracker reminders".');
  H.endBody(html, 'Automatic alert from Mirage AEC Tracker Reminder.');
  await sender.send({
    to: [{ name: 'Tracker Reminder admin', email: cfg.alertEmail }], replyTo: [],
    subject: `[ALERT] Tracker Reminder - ${problems.length} problem(s) - ${H.date(today, today)}`,
    html: html.join(''), inlineImages: [],
  }, null);
}

/** assets/mirage_logo.jpg, or null (emails then go without the logo). */
function readLogo() {
  try {
    return readFileSync(join(HERE, '..', 'assets', LOGO.fileName));
  } catch {
    return null;
  }
}

main().catch((e) => {
  console.error(`!! ${e.message}`);
  process.exitCode = 1;
});
