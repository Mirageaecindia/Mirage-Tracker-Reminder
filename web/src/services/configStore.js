// ================================================================
// web/src/services/configStore.js
// ----------------------------------------------------------------
// PURPOSE : The shared settings files in the Automation site, for the page:
//             trackers.json  (read / save)   - which trackers, columns, times
//             admins.json    (read / save)   - who may use this page
//             name_map.json  (read / save)   - tracker name -> email picks
//             sentlog.json   (read only)     - for the preview
//             runstate.json  (read only)     - "last ran" column
//           Only file access (SRP); the rules for the CONTENT are in ../core.
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Every load returns { data, eTag }; every save takes the eTag it read
//       and returns the new one (see graph.js: no silent overwrites when two
//       admins edit at the same time).
//     * A file that is not valid JSON is an ERROR, never "empty" - an empty
//       trackers.json saved over a broken one would delete every tracker.
//     * Saved JSON is pretty-printed (2 spaces), same as the job, so a human
//       can still read / fix it in VS Code.
//     * The page never writes sentlog.json / runstate.json - only the job does.
// ================================================================

import { SETTINGS } from '../authConfig.js';

export const FILES = Object.freeze({
  trackers: 'trackers.json',
  admins: 'admins.json',
  nameMap: 'name_map.json',
  sentLog: 'sentlog.json',
  runState: 'runstate.json',
});

const WRITABLE = new Set([FILES.trackers, FILES.admins, FILES.nameMap]);

export class ConfigStore {
  /** @param {import('./graph.js').Graph} graph */
  constructor(graph) {
    this.graph = graph;
  }

  path(name) {
    return `${SETTINGS.configFolder}/${name}`;
  }

  /** { data, eTag } - data null when the file does not exist yet. */
  async load(name) {
    const { text, eTag } = await this.graph.readText(SETTINGS.configSiteHost, SETTINGS.configSitePath, this.path(name));
    if (text === null) return { data: null, eTag: null };
    try {
      return { data: JSON.parse(text), eTag };
    } catch (e) {
      throw new Error(`${name} in SharePoint is not valid JSON (${e.message}). Fix it in VS Code or re-upload a good copy before saving here.`);
    }
  }

  /** Saves data; returns the new eTag. Throws ConflictError if someone else saved first. */
  async save(name, data, eTag) {
    if (!WRITABLE.has(name)) throw new Error(`${name} is written by the nightly job only.`);
    return this.graph.writeText(SETTINGS.configSiteHost, SETTINGS.configSitePath, this.path(name), JSON.stringify(data, null, 2), eTag);
  }
}

/** admins.json data -> lower-case email set (missing / malformed -> empty). */
export function adminEmails(data) {
  const list = data && Array.isArray(data.admins) ? data.admins : [];
  return new Set(list.map((a) => String(a && a.email || '').trim().toLowerCase()).filter(Boolean));
}
