// ================================================================
// web/src/services/graph.js
// ----------------------------------------------------------------
// PURPOSE : Microsoft Graph calls AS THE SIGNED-IN ADMIN (delegated):
//           token, requests with time limit + retry, file read/write with
//           version tags, company user list. Only HTTP + auth (SRP).
//           Browser twin of job/graphClient.js (same retry / timeout rules).
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Token: acquireTokenSilent first; if Microsoft needs the user again
//       (expired session, new consent) -> acquireTokenRedirect (redirect, not
//       pop-up: pop-ups are blocked in the Excel side panel later).
//     * 60 s time limit per request + retry on 429 / 503 / 504 / timeout
//       (2 s, 4 s, 8 s) - same rules as the job.
//     * Version tags (eTag): every JSON file is read WITH its eTag and saved
//       with "If-Match: <eTag>". If another admin saved in between,
//       SharePoint answers 412 and we throw ConflictError -> the page says
//       "changed by someone else, reload" instead of silently overwriting.
//     * Users: /users with displayName, mail, userPrincipalName (what
//       User.ReadBasic.All returns). accountEnabled is NOT available with
//       that permission, so the page cannot hide disabled accounts - the job
//       (User.Read.All) still never emails them.
// ================================================================

import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { loginRequest } from '../authConfig.js';

const GRAPH = 'https://graph.microsoft.com/v1.0/';
const RETRY_STATUS = new Set([429, 503, 504]);
const MAX_ATTEMPTS = 4;
const TIMEOUT_MS = 60000;

export class GraphError extends Error {
  constructor(status, code, message, path) {
    super(`Microsoft 365 answered ${status}${code ? ' ' + code : ''} on ${path}: ${message}`);
    this.status = status;
    this.code = code;
  }
}

/** Someone else saved the same file since we read it (see change log). */
export class ConflictError extends Error {}

export class Graph {
  /** @param {import('@azure/msal-browser').IPublicClientApplication} msal */
  constructor(msal) {
    this.msal = msal;
    this.siteIds = new Map();
  }

  async token() {
    const account = this.msal.getActiveAccount();
    try {
      const result = await this.msal.acquireTokenSilent({ ...loginRequest, account });
      return result.accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        await this.msal.acquireTokenRedirect({ ...loginRequest, account });
        return new Promise(() => {}); // page is leaving for the sign-in screen
      }
      throw e;
    }
  }

  /** One request with time limit + retries. path = relative to v1.0, or absolute (paging). */
  async request(method, path, { body, contentType, headers: extra } = {}) {
    const url = /^https:\/\//i.test(path) ? path : GRAPH + path;
    for (let attempt = 1; ; attempt++) {
      const headers = { Authorization: `Bearer ${await this.token()}`, ...(extra || {}) };
      if (contentType) headers['Content-Type'] = contentType;
      let res;
      try {
        res = await fetch(url, { method, headers, body, signal: AbortSignal.timeout(TIMEOUT_MS) });
      } catch (e) {
        if ((e.name === 'TimeoutError' || e.name === 'AbortError') && attempt < MAX_ATTEMPTS) {
          await sleep(2000 * 2 ** (attempt - 1));
          continue;
        }
        throw e.name === 'TimeoutError' ? new Error(`Timed out after 60 s: ${method} ${path.split('?')[0]}`) : e;
      }
      if (res.ok) return res;
      if (res.status === 412) throw new ConflictError('This file was changed by someone else after you opened it. Reload and make your change again.');
      if (RETRY_STATUS.has(res.status) && attempt < MAX_ATTEMPTS) {
        const retryAfter = Number(res.headers.get('retry-after'));
        await sleep(Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 2000 * 2 ** (attempt - 1));
        continue;
      }
      const err = await res.json().catch(() => ({}));
      throw new GraphError(res.status, err.error && err.error.code, (err.error && err.error.message) || res.statusText, path.split('?')[0]);
    }
  }

  async getJson(path) {
    return (await this.request('GET', path)).json();
  }

  async getBytes(path) {
    return (await this.request('GET', path)).arrayBuffer();
  }

  /** "mirageace.sharepoint.com" + "/sites/X" -> site id (cached). */
  async siteId(host, sitePath) {
    const key = `${host}${sitePath}`.toLowerCase();
    if (!this.siteIds.has(key)) this.siteIds.set(key, (await this.getJson(`sites/${host}:${sitePath}`)).id);
    return this.siteIds.get(key);
  }

  /**
   * Text file in a site's default library -> { text, eTag }; { text: null, eTag: null } if it does not exist.
   */
  async readText(host, sitePath, filePath) {
    const item = `sites/${await this.siteId(host, sitePath)}/drive/root:/${encodePath(filePath)}`;
    let meta;
    try {
      meta = await this.getJson(`${item}?$select=eTag`);
    } catch (e) {
      if (e instanceof GraphError && e.status === 404) return { text: null, eTag: null };
      throw e;
    }
    const text = new TextDecoder('utf-8').decode(await this.getBytes(`${item}:/content`)).replace(/^﻿/, '');
    return { text, eTag: meta.eTag };
  }

  /**
   * Create / replace a text file. eTag given -> only if unchanged since read (ConflictError otherwise).
   * eTag null -> only if the file does NOT exist yet (never overwrite blindly). Returns the new eTag.
   */
  async writeText(host, sitePath, filePath, text, eTag, contentType = 'application/json') {
    const item = `sites/${await this.siteId(host, sitePath)}/drive/root:/${encodePath(filePath)}:/content`;
    const headers = eTag ? { 'If-Match': eTag } : { 'If-None-Match': '*' };
    const res = await this.request('PUT', item, { body: text, contentType, headers });
    return (await res.json()).eTag;
  }

  /** Company directory rows (displayName, mail, userPrincipalName), all pages. */
  async listUsers() {
    const users = [];
    let next = 'users?$select=displayName,mail,userPrincipalName&$top=999';
    while (next) {
      const page = await this.getJson(next);
      users.push(...(page.value || []));
      next = page['@odata.nextLink'] || null;
    }
    return users;
  }

  /** Signed-in person { displayName, mail, userPrincipalName }. */
  async me() {
    return this.getJson('me?$select=displayName,mail,userPrincipalName');
  }
}

/** "General/BIM LEADS/a b.xlsx" -> "General/BIM%20LEADS/a%20b.xlsx" */
export function encodePath(filePath) {
  return filePath.replace(/^\/+/, '').split('/').map(encodeURIComponent).join('/');
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
