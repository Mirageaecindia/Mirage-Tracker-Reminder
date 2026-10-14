// ================================================================
// job/graphClient.js
// ----------------------------------------------------------------
// PURPOSE : Talk to Microsoft Graph AS THE JOB APP (no person signed in):
//             - get an app token with the client secret (client-credentials)
//             - send requests, retrying throttling / temporary server errors
//           Only HTTP + auth (SRP). What to call is decided by
//           sharePointFiles.js / directoryService.js / mailSender.js.
//           Equivalent of Excel_Notify\Services\GraphHttp.cs (retry rules
//           kept), but app-only instead of a signed-in user.
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Plain fetch (built into Node 20+), no Graph SDK. WHY (same reason as
//       Excel_Notify): a handful of endpoints does not justify the SDK's
//       dependency tree.
//     * Token cached until 60 s before expiry; one run is a few minutes, so
//       usually one token per night.
//     * Retries: 429 (throttled), 503, 504 - up to 4 attempts, waiting the
//       Retry-After header if Graph sends one, else 2 s, 4 s, 8 s.
//       Other errors fail at once with Graph's own error text.
//     * Downloads: Graph answers /content with a redirect to a pre-signed
//       URL; fetch follows it and drops the Authorization header on the
//       other host, which is what that URL expects.
//     * The client secret is never logged, never put in an error message.
// - 2026-10-14 : ADDED a 60-second time limit on EVERY request (token too).
//     * WHY: GitHub run #4 hung for 6+ minutes after its first log line with
//       no error - a request that never answered. fetch() has no time limit
//       of its own, so the job just waited until GitHub's 15-minute kill.
//     * A timed-out request is retried like a 503 (2 s, 4 s, 8 s); after the
//       last attempt it fails with "Timed out after 60 s: GET <path>", which
//       reaches the alert email instead of silence.
//     * The limit also covers reading the response body (same signal).
// ================================================================

const GRAPH = 'https://graph.microsoft.com/v1.0/';
const RETRY_STATUS = new Set([429, 503, 504]);
const MAX_ATTEMPTS = 4;
const TIMEOUT_MS = 60000; // per request, see change log

export class GraphError extends Error {
  constructor(status, code, message, path) {
    super(`Graph ${status}${code ? ' ' + code : ''} on ${path}: ${message}`);
    this.status = status;
    this.code = code;
  }
}

export class GraphClient {
  /** @param {{tenantId: string, clientId: string, clientSecret: string}} cfg */
  constructor(cfg, fetchImpl = fetch) {
    this.cfg = cfg;
    this.fetch = fetchImpl;
    this.cached = null; // { token, expiresAt }
  }

  /** App token (cached). */
  async token() {
    if (this.cached && Date.now() < this.cached.expiresAt) return this.cached.token;
    const body = new URLSearchParams({
      client_id: this.cfg.clientId,
      client_secret: this.cfg.clientSecret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    });
    const res = await this.fetch(`https://login.microsoftonline.com/${this.cfg.tenantId}/oauth2/v2.0/token`, {
      method: 'POST', body, signal: AbortSignal.timeout(TIMEOUT_MS),
    }).catch((e) => {
      throw isTimeout(e) ? new Error('Timed out after 60 s: sign-in as the job app') : e;
    });
    const json = await res.json().catch(() => ({}));
    if (!res.ok) {
      // e.g. AADSTS7000222 = secret expired -> the alert email says exactly that
      throw new Error(`Sign-in as the job app failed (${res.status}): ${json.error_description || json.error || 'no details'}`);
    }
    this.cached = { token: json.access_token, expiresAt: Date.now() + (json.expires_in - 60) * 1000 };
    return this.cached.token;
  }

  /**
   * One request with retries. path = relative to v1.0 ("sites/...") or an absolute URL (paging links).
   * @returns {Promise<Response>}
   */
  async request(method, path, { body, contentType } = {}) {
    const url = /^https:\/\//i.test(path) ? path : GRAPH + path;
    for (let attempt = 1; ; attempt++) {
      const headers = { Authorization: `Bearer ${await this.token()}` };
      if (contentType) headers['Content-Type'] = contentType;
      let res;
      try {
        res = await this.fetch(url, { method, headers, body, signal: AbortSignal.timeout(TIMEOUT_MS) });
      } catch (e) {
        if (!isTimeout(e)) throw e;
        if (attempt < MAX_ATTEMPTS) {
          await sleep(2000 * 2 ** (attempt - 1));
          continue;
        }
        throw new Error(`Timed out after 60 s (${MAX_ATTEMPTS} attempts): ${method} ${path.split('?')[0]}`);
      }
      if (res.ok) return res;

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

  /** File bytes as a Node Buffer. */
  async getBytes(path) {
    return Buffer.from(await (await this.request('GET', path)).arrayBuffer());
  }

  async putBytes(path, bytes, contentType) {
    return (await this.request('PUT', path, { body: bytes, contentType })).json();
  }

  /** POST JSON; returns nothing (sendMail answers 202 with no body). */
  async postJson(path, data) {
    await this.request('POST', path, { body: JSON.stringify(data), contentType: 'application/json' });
  }
}

/** fetch aborted by AbortSignal.timeout (name differs by Node version). */
function isTimeout(e) {
  return e && (e.name === 'TimeoutError' || e.name === 'AbortError');
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
