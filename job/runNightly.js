// ================================================================
// job/runNightly.js
// ----------------------------------------------------------------
// PURPOSE : The 8 PM run. ONLY orchestrates (like Excel_Notify's
//           NotifyCommand): every decision lives in core\, every Graph call
//           in job\*.js.
//
// FLOW    : 1. settings (job/config.js), app sign-in
//           2. from the Automation site: trackers.json, name_map.json,
//              sentlog.json; company directory from Microsoft 365
//           3. for EACH enabled tracker (one failing never stops the others):
//                download -> snapshot -> layouts -> read (lead's setup)
//                -> People sheet -> resolve names -> decide (due / overdue)
//                -> send each reminder -> record in sent log
//                -> lead summary (only when there is something to report)
//           4. save sentlog.json once, at the end
//           5. anything went wrong -> alert email to ALERT_EMAIL and exit
//              code 1 (GitHub then also marks the run red and emails the
//              repo owner - two independent warnings)
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

const HERE = dirname(fileURLToPath(import.meta.url));

async function main() {
  const cfg = loadJobConfig(); // throws on missing settings -> caught below (no mail possible yet)
  const graph = new GraphClient(cfg);
  const files = new SharePointFiles(graph);
  const configPath = (name) => posix.join(cfg.configFolder, name);
  const today = indiaToday();
  const nowIso = new Date().toISOString();
  const problems = []; // everything that goes into the alert email
  const sender = new MailSender(graph, { senderMailbox: cfg.senderMailbox, logoBytes: readLogo(), dryRun: cfg.dryRun });

  console.log(`Tracker Reminder run - India date ${today}${cfg.dryRun ? ' - DRY RUN' : ''}${cfg.globalTestMode ? ' - GLOBAL TEST MODE' : ''}`);

  try {
    // ---- 2. shared data ----
    const trackersFile = await files.readJson(cfg.configSiteHost, cfg.configSitePath, configPath(CONFIG_FILES.trackers));
    const { trackers, problems: configProblems } = parseTrackersFile(trackersFile);
    problems.push(...configProblems);

    const nameMapFile = await files.readJson(cfg.configSiteHost, cfg.configSitePath, configPath(CONFIG_FILES.nameMap));
    const { map: sharedMap, rejected } = parseNameMap(nameMapFile || []);
    if (rejected) problems.push(`name_map.json: ${rejected} invalid entr${rejected === 1 ? 'y' : 'ies'} ignored (each needs Name + @mirageaec.com Email).`);

    const sentLog = new SentLog(await files.readJson(cfg.configSiteHost, cfg.configSitePath, configPath(CONFIG_FILES.sentLog)));
    const directory = await getCompanyDirectory(graph);
    console.log(`Loaded ${trackers.length} tracker(s), ${sharedMap.size} mapped name(s), ${directory.length} company user(s).`);

    // ---- 3. each tracker ----
    for (const tracker of trackers.filter((t) => t.enabled)) {
      try {
        await runTracker({ tracker, cfg, files, sender, sentLog, sharedMap, directory, today, nowIso });
      } catch (e) {
        problems.push(`${tracker.name}: ${e.message}`);
        console.error(`!! ${tracker.name}: ${e.message}`);
      }
    }

    // ---- 4. save the sent log once ----
    const pruned = sentLog.prune(nowIso);
    if (!cfg.dryRun) await files.writeJson(cfg.configSiteHost, cfg.configSitePath, configPath(CONFIG_FILES.sentLog), sentLog.toJSON());
    console.log(`Sent log: ${sentLog.entries.length} record(s), ${pruned} old record(s) pruned${cfg.dryRun ? ' (not saved - dry run)' : ''}.`);
  } catch (e) {
    problems.push(`Run stopped: ${e.message}`);
    console.error(`!! Run stopped: ${e.message}`);
  }

  // ---- 5. alert ----
  if (problems.length) {
    await sendAlert(sender, cfg, today, problems).catch((e) => console.error(`!! Alert email failed too: ${e.message}`));
    process.exitCode = 1;
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

/** Today's date in India (UTC+05:30) as YYYY-MM-DD. */
function indiaToday() {
  return new Date(Date.now() + 5.5 * 3600 * 1000).toISOString().slice(0, 10);
}

main().catch((e) => {
  console.error(`!! ${e.message}`);
  process.exitCode = 1;
});
