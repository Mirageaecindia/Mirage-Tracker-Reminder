// ================================================================
// job/config.js
// ----------------------------------------------------------------
// PURPOSE : Read the nightly job's settings from environment variables, in
//           ONE place (SRP), and refuse to start if a required one is
//           missing. No other file reads process.env.
//
// WHERE THEY COME FROM:
//   GitHub Actions : repository Secrets (CLIENT_SECRET) and Variables
//                    (everything else) -> see .github/workflows/nightly.yml
//   Local test     : a .env file (never committed - see .gitignore), run with
//                    node --env-file=.env job/runNightly.js
//
//   TENANT_ID          Entra Directory (tenant) ID
//   CLIENT_ID          Job app "Mirage Tracker Reminder - Job" client ID
//   CLIENT_SECRET      Job app secret VALUE                       (SECRET)
//   SENDER_MAILBOX     project-reminders@mirageaec.com
//   CONFIG_SITE_HOST   mirageace.sharepoint.com
//   CONFIG_SITE_PATH   /sites/Automation
//   CONFIG_FOLDER      Application_Files/TrackerReminder
//   ALERT_EMAIL        who gets "the job failed" emails
//   GLOBAL_TEST_MODE   "true" (default) = EVERY tracker in test mode,
//                      whatever trackers.json says. "false" = per tracker.
//   TEST_RECIPIENT     where test-mode mail goes when GLOBAL_TEST_MODE forces it
//   DRY_RUN            "true" = read + decide + print, send NOTHING, write NOTHING
//   IGNORE_SEND_TIME   "true" = run EVERY tracker now, whatever its send time
//                      (manual / test runs). Default "false".
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * No IDs or addresses hard-coded: the repo is PUBLIC, and changing a
//       mailbox must not need a code change. IDs are not secrets, but they
//       still belong in settings, not code.
//     * GLOBAL_TEST_MODE defaults to TRUE. WHY: the safe failure is "mail
//       went to the tester", never "mail went to the whole team by mistake".
//       Going live is a deliberate switch of one variable.
// - 2026-10-14 : ADDED IGNORE_SEND_TIME (per-tracker send times arrived;
//       a manual test run must still be able to run everything at once).
//       Default false = scheduled behaviour.
// ================================================================

const REQUIRED = [
  'TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'SENDER_MAILBOX',
  'CONFIG_SITE_HOST', 'CONFIG_SITE_PATH', 'CONFIG_FOLDER', 'ALERT_EMAIL', 'TEST_RECIPIENT',
];

/** Settings object; throws listing EVERY missing variable at once. */
export function loadJobConfig(env = process.env) {
  const missing = REQUIRED.filter((k) => !String(env[k] || '').trim());
  if (missing.length) throw new Error(`Missing settings: ${missing.join(', ')}. See job/config.js.`);

  return Object.freeze({
    tenantId: env.TENANT_ID.trim(),
    clientId: env.CLIENT_ID.trim(),
    clientSecret: env.CLIENT_SECRET.trim(),
    senderMailbox: env.SENDER_MAILBOX.trim(),
    configSiteHost: env.CONFIG_SITE_HOST.trim(),
    configSitePath: env.CONFIG_SITE_PATH.trim(),
    configFolder: env.CONFIG_FOLDER.trim().replace(/^\/+|\/+$/g, ''),
    alertEmail: env.ALERT_EMAIL.trim(),
    testRecipient: env.TEST_RECIPIENT.trim(),
    globalTestMode: String(env.GLOBAL_TEST_MODE ?? 'true').trim().toLowerCase() !== 'false', // default ON
    dryRun: String(env.DRY_RUN ?? 'false').trim().toLowerCase() === 'true',
    ignoreSendTime: String(env.IGNORE_SEND_TIME ?? 'false').trim().toLowerCase() === 'true',
  });
}
