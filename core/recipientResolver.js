// ================================================================
// recipientResolver.js
// ----------------------------------------------------------------
// PURPOSE : Match a tracker name ("Abhi", "Ruwaisid") to ONE company user.
//           Only matching (SRP). Reading the directory / name map from
//           Microsoft 365 / SharePoint is the job's I/O code.
//           MIRRORS Excel_Notify\Services\RecipientResolver.cs (+ the
//           directory filter of GraphDirectoryService.cs and the file format
//           of NameMapStore.cs) - keep them in step.
//
// ORDER   : 1. Name maps, in priority order (user decision 2026-10-13):
//                a) SHARED name map in SharePoint (picks made on the admin page)
//                b) the tracker's own "People" sheet (peopleSheetReader.js)
//              A mapped address is used only if that person still exists in
//              the directory (left the company -> fall through).
//           2. Exact display-name match            ("Santhosh Gowda")
//           3. Exact match on the address part     ("santhosh.gowda")
//           Otherwise: NOT matched (null). Never a guess.
//
// CHANGE LOG
// - 2026-10-14 : Initial version, ported from RecipientResolver.cs.
//     * Names normalised before comparing: lower case, trimmed, '.', '_' and
//       '-' treated as spaces, repeated spaces collapsed (same as C#).
//     * A name that matches TWO OR MORE users is NOT matched (same as C#):
//       picking one of two "Aakash" would email the wrong person.
//     * No fuzzy matching ("santosh" vs "santhosh") on purpose - the lead
//       maps it once on the admin page and the shared map remembers it.
//     * CHANGED from C#: several name maps instead of one (shared map +
//       People sheet). Map keys compared trimmed + case-insensitive, like
//       C# NameMapStore (OrdinalIgnoreCase + Trim).
//     * ALLOWED_DOMAIN: mail only ever goes to @mirageaec.com (tenant has
//       guest accounts; a wrong match must never send project data outside
//       the company). Checked here AND again right before sending.
//     * parseNameMap() reads Excel_Notify's export format
//       [ { "Name": "...", "Email": "..." } ], so an existing Excel Notifier
//       name map can be imported into the shared map as-is.
// ================================================================

/** Only addresses ending in "@" + this domain may receive mail. */
export const ALLOWED_DOMAIN = 'mirageaec.com';

/** True if the address is a company address (see change log). */
export function isAllowedAddress(email) {
  if (typeof email !== 'string') return false;
  const pattern = new RegExp(`^[^@\\s]+@${ALLOWED_DOMAIN.replace(/\./g, '\\.')}$`, 'i');
  return pattern.test(email.trim());
}

/** Name-map key: trimmed, case-insensitive (C# NameMapStore.Key + OrdinalIgnoreCase). */
export function nameMapKey(name) {
  return String(name ?? '').trim().toLowerCase();
}

/**
 * Excel_Notify name-map JSON -> Map(key -> email). Invalid entries are skipped and counted.
 * @returns {{map: Map<string,string>, rejected: number}}
 */
export function parseNameMap(entries) {
  const map = new Map();
  let rejected = 0;
  for (const e of Array.isArray(entries) ? entries : []) {
    if (!e || !String(e.Name ?? '').trim() || !isAllowedAddress(e.Email)) {
      rejected++;
      continue;
    }
    map.set(nameMapKey(e.Name), String(e.Email).trim());
  }
  return { map, rejected };
}

/** Map -> Excel_Notify name-map JSON (sorted by name), for saving. */
export function nameMapToEntries(map) {
  return [...map.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([Name, Email]) => ({ Name, Email }));
}

/**
 * Graph /users rows -> company users (mirrors GraphDirectoryService.AddIfCompany):
 * address = mail, or userPrincipalName when mail is empty; only @mirageaec.com;
 * sorted by display name.
 * @param {{displayName?: string, mail?: string, userPrincipalName?: string, accountEnabled?: boolean}[]} users
 */
export function directoryFromGraphUsers(users) {
  const result = [];
  for (const u of users || []) {
    if (u.accountEnabled === false) continue; // NEW vs C#: left staff (disabled) are never emailed
    const address = String((u.mail && u.mail.trim()) || u.userPrincipalName || '').trim();
    if (!address || !isAllowedAddress(address)) continue; // guest / external account
    result.push({ displayName: (u.displayName || '').trim() || address, email: address });
  }
  return result.sort((a, b) => a.displayName.localeCompare(b.displayName, undefined, { sensitivity: 'base' }));
}

export class RecipientResolver {
  /**
   * @param {{displayName: string, email: string}[]} directory  company users
   * @param {Map<string,string>[]} nameMaps  in priority order (shared map first)
   */
  constructor(directory, nameMaps = []) {
    this.directory = directory;
    this.nameMaps = nameMaps;
  }

  /** The matched user {displayName, email}, or null when there is no single sure match. */
  resolve(trackerName) {
    if (trackerName == null || String(trackerName).trim() === '') return null;

    // 1. Name maps (shared map, then People sheet)
    const key = nameMapKey(trackerName);
    for (const map of this.nameMaps) {
      const mapped = map.get(key);
      if (!mapped) continue;
      const user = this.directory.find((u) => u.email.toLowerCase() === mapped.toLowerCase());
      if (user) return user; // else: person left the company -> next source
    }

    const wanted = normalise(trackerName);

    // 2. Display name
    const byName = single(this.directory.filter((u) => normalise(u.displayName) === wanted));
    if (byName) return byName;

    // 3. Address part before '@'
    return single(this.directory.filter((u) => normalise(localPart(u.email)) === wanted));
  }
}

/** Exactly one match, otherwise null (ambiguous = not matched). */
function single(matches) {
  return matches.length === 1 ? matches[0] : null;
}

function localPart(email) {
  const at = email.indexOf('@');
  return at > 0 ? email.slice(0, at) : email;
}

function normalise(text) {
  return String(text ?? '').toLowerCase().replace(/[._-]/g, ' ').replace(/\s+/g, ' ').trim();
}
