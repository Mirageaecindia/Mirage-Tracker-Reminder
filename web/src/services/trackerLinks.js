// ================================================================
// web/src/services/trackerLinks.js
// ----------------------------------------------------------------
// PURPOSE : Turn ANY SharePoint link a lead pastes into the tracker's
//           location { siteHost, sitePath, filePath, fileName }, and download
//           the workbook. Only link handling + download (SRP).
//
// LINKS HANDLED:
//   1. Address-bar file link   https://<host>/sites/X/Shared%20Documents/a/b.xlsx?web=1
//   2. Library view link       https://<host>/sites/X/Shared%20Documents/Forms/AllItems.aspx?id=/sites/X/Shared Documents/a/b.xlsx
//   3. "Copy link" sharing link https://<host>/:x:/s/X/EabcXYZ...?e=...   (and /:x:/r/... links)
//      -> resolved through Graph /shares (the admin's own access decides).
//
// CHANGE LOG
// - 2026-10-14 : Initial version (found while planning: most people paste
//       "Copy link" sharing links, which core/trackerConfig.parseSharePointUrl
//       cannot read on its own).
//     * Same rules as parseSharePointUrl: OneDrive files are refused (they
//       break when the owner leaves); only the site's MAIN document library
//       is accepted (the job reads files by path in the default library).
//     * Only .xlsx / .xlsm.
// ================================================================

import { parseSharePointUrl } from '../../../core/trackerConfig.js';
import { encodePath } from './graph.js';

/** Any supported link -> { siteHost, sitePath, filePath, fileName }. Throws a readable Error. */
export async function resolveTrackerLink(graph, link) {
  const text = String(link || '').trim();
  let url;
  try {
    url = new URL(text);
  } catch {
    throw new Error('That is not a link. Copy it from the browser address bar or from SharePoint "Copy link".');
  }
  if (/-my\.sharepoint\.com$/i.test(url.hostname)) {
    throw new Error('This file is in someone\'s OneDrive. Move it to a team site first (it would break when that person leaves).');
  }
  if (!/\.sharepoint\.com$/i.test(url.hostname)) throw new Error('Not a SharePoint link.');

  // 2. library view link: the real path is in ?id=
  const idParam = url.searchParams.get('id');
  if (idParam && /^\/(sites|teams)\//i.test(idParam)) {
    return withName(parseSharePointUrl(`https://${url.hostname}${encodeURI(idParam)}`));
  }

  // 1. plain file link
  if (/^\/(sites|teams)\//i.test(url.pathname)) {
    return withName(parseSharePointUrl(text));
  }

  // 3. sharing link
  return resolveSharingLink(graph, text, url.hostname);
}

/** Workbook bytes for a resolved location (as the signed-in admin). */
export async function downloadTracker(graph, loc) {
  const siteId = await graph.siteId(loc.siteHost, loc.sitePath);
  return graph.getBytes(`sites/${siteId}/drive/root:/${encodePath(loc.filePath)}:/content`);
}

async function resolveSharingLink(graph, link, host) {
  const shareId = 'u!' + btoa(unescape(encodeURIComponent(link))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  let item;
  try {
    item = await graph.getJson(`shares/${shareId}/driveItem?$select=name,parentReference`);
  } catch (e) {
    throw new Error(`SharePoint could not open that link with your account (${e.message}).`);
  }
  const parent = item.parentReference || {};
  if (!parent.siteId || !parent.driveId || !parent.path) throw new Error('That link does not point to a file in a team site.');

  const site = await graph.getJson(`sites/${parent.siteId}?$select=webUrl`);
  const sitePath = new URL(site.webUrl).pathname.replace(/\/+$/, '');
  if (!/^\/(sites|teams)\/[^/]+$/i.test(sitePath)) throw new Error('That link does not point to a file in a team site.');

  const mainDrive = await graph.getJson(`sites/${parent.siteId}/drive?$select=id`);
  if (mainDrive.id !== parent.driveId) {
    throw new Error('Only the site\'s main "Documents" library is supported - this file is in another library.');
  }

  const folder = decodeURIComponent(parent.path.replace(/^.*?root:/, '')).replace(/^\/+/, '');
  return withName({ siteHost: host, sitePath, filePath: folder ? `${folder}/${item.name}` : item.name });
}

function withName(loc) {
  const fileName = loc.filePath.split('/').pop();
  if (!/\.xls[xm]$/i.test(fileName)) throw new Error(`"${fileName}" is not an Excel workbook (.xlsx / .xlsm).`);
  return { ...loc, fileName };
}
