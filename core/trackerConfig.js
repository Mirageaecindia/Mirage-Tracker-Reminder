// ================================================================
// trackerConfig.js
// ----------------------------------------------------------------
// PURPOSE : The format of trackers.json (the list of registered trackers)
//           and its validation. Pure data rules (SRP); reading / writing the
//           file in SharePoint is the job's / web page's I/O code.
//           Replaces Excel_Notify's per-PC SheetSetupStore: the setup is now
//           shared, because the nightly job must see what the lead chose.
//
// FILE    : Automation site -> Documents -> Application_Files ->
//           TrackerReminder -> trackers.json
//   {
//     "version": 1,
//     "trackers": [ {
//       "id": "test-mep",                       // stable, a-z 0-9 -
//       "name": "TEST Tracker - MEP",           // shown in emails
//       "enabled": true,
//       "testMode": true,                       // all mail -> testRecipient
//       "testRecipient": "you@mirageaec.com",
//       "siteHost": "mirageace.sharepoint.com",
//       "sitePath": "/sites/Automation",
//       "filePath": "Application_Files/TrackerReminder/test/TEST Tracker - MEP.xlsx",
//       "lead": { "name": "...", "email": "...@mirageaec.com" },
//       "senderMode": "default" | "onBehalf",
//       "onBehalfOf": { "name": "...", "email": "..." } | null,
//       "backlogPolicy": "afterRegistration" | "emailOnce" | "leadSummaryOnly",
//       "resendPolicy": "once" | "daily" | "everyRun",   // default "once"
//       "coLeads": [ { "name": "...", "email": "...@mirageaec.com" } ],   // max 3, also get the summary
//       "ccReminders": [ "a@mirageaec.com" ],   // CC on every reminder email
//       "ccSummary": [ "b@mirageaec.com" ],     // CC on the lead summary
//       "registeredOn": "2026-10-14",
//       "sendTime": "20:00",                     // India time, HH:00 or HH:30
//       "sheets": [ { "name": "F-14",
//                     "roles": { "Engineers": "Engineering Inputs",
//                                "Modelers": "Modeller With Coordination" } } ]
//     } ]
//   }
//   filePath is relative to the site's DEFAULT document library
//   ("Documents" / "Shared Documents").
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Roles are saved by enum NAME and blocks by column TITLE - same
//       choices as Excel_Notify's SheetSetup (title survives inserted
//       columns; a vanished title becomes a warning, not a wrong column).
//     * A tracker with ANY validation problem is skipped and the problem is
//       reported - never half-run with guessed settings.
//     * testMode defaults to TRUE when missing (safe side).
//     * parseSharePointUrl(): turns the link the lead pastes into
//       siteHost / sitePath / filePath. Only the default library
//       ("Shared Documents") is supported for now - anything else is
//       rejected with a clear message.
// - 2026-10-14 : ADDED sendTime (user decision: per-tracker send time in
//       IST, default 8 PM, 30-minute steps). Missing -> "20:00"; anything
//       other than HH:00 / HH:30 is a validation problem (see schedule.js
//       for why only half hours).
// - 2026-10-15 : ADDED resendPolicy, coLeads, ccReminders, ccSummary (user request).
//     WHY : leads wanted (1) more than one lead on the summary, (2) CC lists for
//           reminders and for the summary, (3) a choice between "send once" and
//           "keep resending until fixed".
//     WHAT: new optional fields with safe defaults (once / [] / [] / []), so a
//           trackers.json saved before this change loads and behaves as before.
//           Validation: coLeads max 3 and company addresses only; CC lists
//           company addresses only (same rule as To / Reply-To); resendPolicy
//           must be one of ResendPolicy. Test mode: CC is dropped by mailSender.js.
// ================================================================

import { roleFromName } from './roles.js';
import { SheetSetup } from './assignmentItem.js';
import { isAllowedAddress } from './recipientResolver.js';
import { BacklogPolicy, ResendPolicy } from './recipientBuilder.js';
import { SenderMode } from './emailComposer.js';
import { DEFAULT_SEND_TIME, SEND_TIME } from './schedule.js';

const ID = /^[a-z0-9][a-z0-9-]{1,40}$/;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
export const MAX_CO_LEADS = 3;

/**
 * Parsed trackers.json -> valid trackers + problems.
 * @returns {{trackers: object[], problems: string[]}}
 */
export function parseTrackersFile(data) {
  const result = { trackers: [], problems: [] };
  if (!data || !Array.isArray(data.trackers)) {
    result.problems.push('trackers.json is missing or has no "trackers" list.');
    return result;
  }
  const seen = new Set();
  data.trackers.forEach((raw, index) => {
    const t = withDefaults(raw);
    const problems = validateTracker(t);
    if (t.id && seen.has(t.id)) problems.push(`duplicate id "${t.id}"`);
    seen.add(t.id);
    if (problems.length) result.problems.push(`Tracker #${index + 1} (${t.name || t.id || 'no name'}) skipped: ${problems.join('; ')}.`);
    else result.trackers.push(t);
  });
  return result;
}

/** Missing optional fields -> safe defaults (see change log). */
export function withDefaults(raw) {
  const t = { ...(raw || {}) };
  t.enabled = t.enabled !== false;
  t.testMode = t.testMode !== false;
  t.senderMode = t.senderMode || SenderMode.Default;
  t.backlogPolicy = t.backlogPolicy || BacklogPolicy.AfterRegistration;
  t.onBehalfOf = t.onBehalfOf || null;
  t.resendPolicy = t.resendPolicy || ResendPolicy.Once;
  t.coLeads = Array.isArray(t.coLeads) ? t.coLeads : [];
  t.ccReminders = Array.isArray(t.ccReminders) ? t.ccReminders : [];
  t.ccSummary = Array.isArray(t.ccSummary) ? t.ccSummary : [];
  t.sendTime = t.sendTime || DEFAULT_SEND_TIME;
  t.sheets = Array.isArray(t.sheets) ? t.sheets : [];
  return t;
}

/** List of problems; empty = valid. */
export function validateTracker(t) {
  const p = [];
  if (!ID.test(t.id || '')) p.push('id must be 2-41 chars a-z 0-9 -');
  if (!String(t.name || '').trim()) p.push('name missing');
  if (!/^[a-z0-9.-]+\.sharepoint\.com$/i.test(t.siteHost || '')) p.push('siteHost must be <tenant>.sharepoint.com');
  if (!/^\/(sites|teams)\/[^/]+$/i.test(t.sitePath || '')) p.push('sitePath must look like /sites/Name');
  if (!/\.xls[xm]$/i.test(t.filePath || '')) p.push('filePath must point to an .xlsx / .xlsm file');
  if (!t.lead || !isAllowedAddress(t.lead.email)) p.push('lead.email must be a @mirageaec.com address');
  if (t.testMode && !isAllowedAddress(t.testRecipient)) p.push('testRecipient must be a @mirageaec.com address while testMode is on');
  if (!Object.values(SenderMode).includes(t.senderMode)) p.push(`senderMode must be ${Object.values(SenderMode).join(' / ')}`);
  if (t.senderMode === SenderMode.OnBehalf && (!t.onBehalfOf || !isAllowedAddress(t.onBehalfOf.email))) {
    p.push('onBehalfOf.email must be a @mirageaec.com address in On-behalf mode');
  }
  if (!Object.values(BacklogPolicy).includes(t.backlogPolicy)) p.push(`backlogPolicy must be ${Object.values(BacklogPolicy).join(' / ')}`);
  if (!Object.values(ResendPolicy).includes(t.resendPolicy)) p.push(`resendPolicy must be ${Object.values(ResendPolicy).join(' / ')}`);
  if (t.coLeads.length > MAX_CO_LEADS) p.push(`at most ${MAX_CO_LEADS} co-leads`);
  t.coLeads.forEach((c) => { if (!c || !isAllowedAddress(c.email)) p.push(`co-lead "${c && c.email}" must be a @mirageaec.com address`); });
  t.ccReminders.forEach((e) => { if (!isAllowedAddress(e)) p.push(`CC (reminders) "${e}" must be a @mirageaec.com address`); });
  t.ccSummary.forEach((e) => { if (!isAllowedAddress(e)) p.push(`CC (summary) "${e}" must be a @mirageaec.com address`); });
  if (!ISO_DATE.test(t.registeredOn || '')) p.push('registeredOn must be YYYY-MM-DD');
  if (!SEND_TIME.test(t.sendTime || '')) p.push('sendTime must be HH:00 or HH:30 (India time, e.g. "20:00")');
  if (t.sheets.length === 0) p.push('no sheets configured');
  t.sheets.forEach((s) => {
    if (!s || !String(s.name || '').trim()) p.push('a sheet has no name');
    else if (!s.roles || Object.keys(s.roles).length === 0) p.push(`sheet "${s.name}" has no roles`);
    else {
      for (const roleName of Object.keys(s.roles)) {
        if (roleFromName(roleName) === undefined) p.push(`sheet "${s.name}": unknown role "${roleName}"`);
        if (!String(s.roles[roleName] || '').trim()) p.push(`sheet "${s.name}": role "${roleName}" has no column title`);
      }
    }
  });
  return p;
}

/** Tracker config -> SheetSetups for trackerReader (role names -> role ids). */
export function toSheetSetups(tracker) {
  return tracker.sheets.map((s) => {
    const setup = new SheetSetup(s.name);
    for (const [roleName, title] of Object.entries(s.roles)) {
      const role = roleFromName(roleName);
      if (role === undefined) continue; // already reported by validateTracker
      setup.roles.add(role);
      setup.blockTitles.set(role, title);
    }
    return setup;
  });
}

/**
 * "https://mirageace.sharepoint.com/sites/X/Shared%20Documents/General/a.xlsx?web=1"
 *   -> { siteHost, sitePath: "/sites/X", filePath: "General/a.xlsx" }
 * Throws with a readable message for anything else (see change log).
 */
export function parseSharePointUrl(url) {
  let u;
  try {
    u = new URL(String(url).trim());
  } catch {
    throw new Error('Not a valid link.');
  }
  if (!/\.sharepoint\.com$/i.test(u.hostname)) throw new Error('Not a SharePoint link.');
  if (/-my\.sharepoint\.com$/i.test(u.hostname)) {
    throw new Error('This file is in someone\'s OneDrive. Move it to a team site first (it would break when that person leaves).');
  }
  const parts = decodeURIComponent(u.pathname).split('/').filter(Boolean); // ["sites","X","Shared Documents",...]
  if (parts.length < 4 || !/^(sites|teams)$/i.test(parts[0])) throw new Error('Link does not point to a file in a team site.');
  if (parts[2].toLowerCase() !== 'shared documents') {
    throw new Error(`Only the site's main "Documents" library is supported (this link is in "${parts[2]}").`);
  }
  return { siteHost: u.hostname, sitePath: `/${parts[0]}/${parts[1]}`, filePath: parts.slice(3).join('/') };
}
