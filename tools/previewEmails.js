// ================================================================
// previewEmails.js
// ----------------------------------------------------------------
// PURPOSE : Offline DRY RUN of tonight's job for one tracker file: reads the
//           tracker, decides who gets what, and WRITES the emails as .html
//           files into out\ so they can be opened in a browser.
//           NOTHING is sent, nothing is written to SharePoint.
//
// USAGE   : node tools/previewEmails.js "<tracker.xlsx>" [options]
//             --today YYYY-MM-DD       pretend "today" (default: India date)
//             --registered YYYY-MM-DD  tracker registration date (default: today)
//             --backlog afterRegistration|emailOnce|leadSummaryOnly
//             --names <file.json>      Excel_Notify name-map export
//                                      [ { "Name": "...", "Email": "..." } ]
//             --demo                   no names file: give EVERY tracker name a
//                                      fake address name@mirageaec.com, just to
//                                      see the email layout
//             --mode default|onBehalf  sender mode
//             --lead "Name <email>"    lead for Reply-To + summary
//           Output: out\<n>-<sheet>-<person>.html and out\lead-summary.html
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Offline = no Microsoft 365 directory. The directory is built from
//       the name-map file (or --demo), so the resolver behaves as in the job
//       for the names it knows.
//     * The roles/columns are each sheet's AUTO-MATCHED blocks (same as
//       parityCheck.js); in the real job they come from the lead's saved
//       setup on the admin page.
//     * cid:mirage_logo is replaced by ../assets/mirage_logo.jpg so the logo
//       shows in a browser (same trick as Excel_Notify's PreviewHtml).
//     * out\ is in .gitignore (contains staff names).
// ================================================================

import { readFileSync, mkdirSync, writeFileSync, readdirSync, unlinkSync } from 'node:fs';
import { basename, join } from 'node:path';
import XLSX from 'xlsx';

import { takeSnapshot } from '../core/workbookSnapshot.js';
import { detectLayout } from '../core/sheetLayoutDetector.js';
import { SheetSetup } from '../core/assignmentItem.js';
import { readTracker } from '../core/trackerReader.js';
import { readPeopleSheet } from '../core/peopleSheetReader.js';
import { RecipientResolver, parseNameMap, nameMapKey } from '../core/recipientResolver.js';
import { buildReminders, BacklogPolicy } from '../core/recipientBuilder.js';
import { SentLog } from '../core/sentLog.js';
import { composeReminder, SenderMode, LOGO } from '../core/emailComposer.js';
import { composeLeadSummary } from '../core/leadSummaryComposer.js';

// ---- arguments ----------------------------------------------------------------
const args = process.argv.slice(2);
const file = args.find((a, i) => !a.startsWith('--') && !(i > 0 && args[i - 1].startsWith('--') && args[i - 1] !== '--demo'));
if (!file) {
  console.error('Usage: node tools/previewEmails.js "<tracker.xlsx>" [--today YYYY-MM-DD] [--names map.json | --demo] ...');
  process.exit(1);
}
const opt = (name, fallback) => (args.includes(name) ? args[args.indexOf(name) + 1] : fallback);
const today = opt('--today', indiaToday());
const registeredOn = opt('--registered', today);
const backlogPolicy = opt('--backlog', BacklogPolicy.AfterRegistration);
const senderMode = opt('--mode', SenderMode.Default);
const lead = parsePerson(opt('--lead', 'Tracker Lead <project-reminders@mirageaec.com>'));
const demo = args.includes('--demo');

// ---- read the tracker (same steps as the job) ------------------------------------
const snapshot = takeSnapshot(XLSX, readFileSync(file), { name: basename(file), fullName: file });
const layouts = new Map(snapshot.sheets.map((s) => [s.name, detectLayout(s)]));
const setups = [];
for (const [name, layout] of layouts) {
  if (!layout.isTracker) continue;
  const setup = new SheetSetup(name);
  for (const [role, block] of layout.autoMatched) {
    setup.roles.add(role);
    setup.blockTitles.set(role, block.title);
  }
  if (setup.roles.size) setups.push(setup);
}
const read = readTracker(snapshot, layouts, setups);
const people = readPeopleSheet(snapshot);

// ---- offline directory + name maps -----------------------------------------------
let sharedMap = new Map();
if (args.includes('--names')) sharedMap = parseNameMap(JSON.parse(readFileSync(opt('--names'), 'utf8'))).map;
if (demo) {
  for (const item of read.items) {
    const fake = item.personName.trim().toLowerCase().replace(/[^a-z0-9]+/g, '.') + '@mirageaec.com';
    if (!sharedMap.has(nameMapKey(item.personName))) sharedMap.set(nameMapKey(item.personName), fake);
  }
}
const directory = [...new Set([...sharedMap.values(), ...people.map.values()])]
  .map((email) => ({ displayName: email.split('@')[0], email }));
const resolver = new RecipientResolver(directory, [sharedMap, people.map]);

// ---- decide + compose ------------------------------------------------------------
const plan = buildReminders({
  items: read.items, today, trackerId: 'preview', registeredOn, backlogPolicy, resolver, sentLog: new SentLog(),
});

const outDir = 'out';
mkdirSync(outDir, { recursive: true });
for (const f of readdirSync(outDir)) if (f.endsWith('.html')) unlinkSync(join(outDir, f)); // fresh preview

const results = [];
plan.recipients.forEach((recipient, i) => {
  const mail = composeReminder(recipient, {
    workbookName: snapshot.name, today, senderMode, lead,
    onBehalfOf: senderMode === SenderMode.OnBehalf ? lead : undefined, includeLogo: true,
  });
  const fileName = `${String(i + 1).padStart(2, '0')}-${safe(recipient.sheetName)}-${safe(recipient.names[0])}.html`;
  writeFileSync(join(outDir, fileName), pageFor(mail));
  results.push({ recipient, ok: true });
});

const summary = composeLeadSummary({
  trackerName: snapshot.name, workbookName: snapshot.name, today, lead, results,
  unmatched: plan.unmatched, backlog: plan.backlog, alreadySentCount: plan.alreadySentCount,
  warnings: [...read.warnings, ...people.warnings], errors: [], testMode: true, includeLogo: true,
});
writeFileSync(join(outDir, 'lead-summary.html'), pageFor(summary));

// ---- console report ----------------------------------------------------------------
console.log(`Tracker      : ${snapshot.name}`);
console.log(`Today        : ${today}   Registered: ${registeredOn}   Backlog: ${backlogPolicy}   Mode: ${senderMode}`);
console.log(`Reminders    : ${plan.recipients.length} email(s)`);
plan.recipients.forEach((r) => {
  const d = r.entries.filter((e) => e.type === 'due').length;
  const o = r.entries.filter((e) => e.type === 'overdue').length;
  console.log(`   ${r.sheetName.padEnd(24)} ${r.names.join(', ').padEnd(24)} ${r.email.padEnd(34)} due ${d}  overdue ${o}`);
});
console.log(`Not matched  : ${plan.unmatched.length}${plan.unmatched.length ? ' -> ' + plan.unmatched.map((u) => `${u.name} (${u.sheetName})`).join(', ') : ''}`);
console.log(`Backlog      : ${plan.backlog.length} listed for lead, ${plan.backlogSkippedCount} skipped`);
console.log(`People sheet : ${people.sheetName || 'none'} (${people.map.size} names)`);
console.log(`Written to   : ${outDir}\\  (open the .html files in a browser)`);

// ------------------------------------------------------------
function pageFor(mail) {
  const head = `<p style="font:12px Segoe UI,Arial;color:#555;border-bottom:1px solid #ccc;padding-bottom:6px;">`
    + `<b>To:</b> ${esc(mail.to.map((t) => `${t.name} <${t.email}>`).join('; '))}<br>`
    + `<b>Reply-To:</b> ${esc(mail.replyTo.map((t) => `${t.name} <${t.email}>`).join('; ') || '-')}<br>`
    + `<b>Subject:</b> ${esc(mail.subject)}</p>`;
  const body = mail.html.split(`cid:${LOGO.contentId}`).join(`../assets/${LOGO.fileName}`);
  return `<!doctype html><meta charset="utf-8"><title>${esc(mail.subject)}</title>${head}${body}`;
}
function esc(t) {
  return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function safe(t) {
  return String(t).replace(/[^A-Za-z0-9]+/g, '_').replace(/^_|_$/g, '').slice(0, 30);
}
function parsePerson(text) {
  const m = /^\s*(.*?)\s*<([^>]+)>\s*$/.exec(text);
  return m ? { name: m[1] || m[2], email: m[2] } : { name: text, email: text };
}
function indiaToday() {
  return new Date(Date.now() + 5.5 * 3600 * 1000).toISOString().slice(0, 10);
}
