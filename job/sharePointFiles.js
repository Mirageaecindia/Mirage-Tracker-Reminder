// ================================================================
// job/sharePointFiles.js
// ----------------------------------------------------------------
// PURPOSE : Files in SharePoint, as the job app:
//             - download a tracker workbook (site + path in the library)
//             - read / write the JSON config files (trackers.json,
//               name_map.json, sentlog.json) in the TrackerReminder folder
//           Only file I/O (SRP).
//
// PERMISSIONS (granted by IT, 2026-10-13, Sites.Selected):
//           Automation site        -> write  (config files, test tracker)
//           MirageAEC_Engg_BIM     -> read   (live trackers)
//           Any other site         -> 403 -> reported as "no access to site"
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Files addressed by PATH in the site's default library
//       (/sites/{id}/drive/root:/{path}:/content). WHY: works with
//       Sites.Selected, and the path is readable in trackers.json. Limit: a
//       renamed / moved tracker is "not found" -> lead summary + alert say so.
//     * Each path segment is URL-encoded separately ("BIM LEADS" ->
//       "BIM%20LEADS"), the slashes are kept.
//     * Site ids looked up once per site per run (cache).
//     * readJson: missing file -> null (first run has no sentlog.json yet);
//       a file that is not valid JSON -> error (never silently "empty",
//       which could re-send every overdue reminder).
// ================================================================

import { GraphError } from './graphClient.js';

export const CONFIG_FILES = Object.freeze({
  trackers: 'trackers.json',
  nameMap: 'name_map.json',
  sentLog: 'sentlog.json',
});

export class SharePointFiles {
  /** @param {import('./graphClient.js').GraphClient} graph */
  constructor(graph) {
    this.graph = graph;
    this.siteIds = new Map();
  }

  /** "mirageace.sharepoint.com" + "/sites/X" -> Graph site id (cached). */
  async siteId(host, sitePath) {
    const key = `${host}${sitePath}`.toLowerCase();
    if (!this.siteIds.has(key)) {
      try {
        const site = await this.graph.getJson(`sites/${host}:${sitePath}`);
        this.siteIds.set(key, site.id);
      } catch (e) {
        if (e instanceof GraphError && (e.status === 403 || e.status === 401)) {
          throw new Error(`No access to site ${sitePath} - IT must grant the job app access to it (Sites.Selected).`);
        }
        throw e;
      }
    }
    return this.siteIds.get(key);
  }

  /** Workbook bytes; clear message when the file is missing. */
  async downloadFile(host, sitePath, filePath) {
    const id = await this.siteId(host, sitePath);
    try {
      return await this.graph.getBytes(`sites/${id}/drive/root:/${encodePath(filePath)}:/content`);
    } catch (e) {
      if (e instanceof GraphError && e.status === 404) {
        throw new Error(`File not found: ${sitePath}/${filePath} (renamed or moved?)`);
      }
      throw e;
    }
  }

  /** Parsed JSON, or null when the file does not exist yet. */
  async readJson(host, sitePath, filePath) {
    let bytes;
    try {
      bytes = await this.downloadFile(host, sitePath, filePath);
    } catch (e) {
      if (/^File not found/.test(e.message)) return null;
      throw e;
    }
    try {
      return JSON.parse(bytes.toString('utf8').replace(/^﻿/, '')); // tolerate a BOM from Notepad
    } catch {
      throw new Error(`${filePath} is not valid JSON - fix or delete it (see change log).`);
    }
  }

  /** Creates or replaces a JSON file (pretty-printed, so a human can read it). */
  async writeJson(host, sitePath, filePath, data) {
    const id = await this.siteId(host, sitePath);
    const body = Buffer.from(JSON.stringify(data, null, 2), 'utf8');
    await this.graph.putBytes(`sites/${id}/drive/root:/${encodePath(filePath)}:/content`, body, 'application/json');
  }
}

/** "General/BIM LEADS/a b.xlsx" -> "General/BIM%20LEADS/a%20b.xlsx" */
function encodePath(filePath) {
  return filePath.replace(/^\/+/, '').split('/').map(encodeURIComponent).join('/');
}
