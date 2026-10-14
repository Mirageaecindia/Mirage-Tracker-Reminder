// ================================================================
// web/src/authConfig.js
// ----------------------------------------------------------------
// PURPOSE : Sign-in + location settings for the admin page, in ONE place.
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * App = "Mirage Tracker Reminder - Admin" (SPA registration, IT
//       2026-10-13). Tenant / client IDs are NOT secrets: a browser app
//       has to send them in every sign-in URL anyway, so they live in code.
//       There is NO secret in the web page, by design.
//     * redirectUri = this site's own address (origin + base path), which is
//       exactly one of the two registered URIs:
//         https://mirageaecindia-dev.github.io/Mirage-Tracker-Reminder/
//         http://localhost:5173/                       (local dev)
//     * Scopes = the delegated permissions IT consented. "Delegated" = the
//       page can only do what the signed-in admin can do in SharePoint.
//     * Token cache in localStorage so a sign-in survives a page reload and
//       other tabs (the Excel side panel later shares nothing with Edge).
// ================================================================

export const TENANT_ID = '37f52e66-51e4-475f-a8ee-cd68a2eb2971';
export const CLIENT_ID = 'fdac89f2-ef66-4c11-bdec-f820d26108d6';

export const SCOPES = ['User.Read', 'User.ReadBasic.All', 'Sites.ReadWrite.All'];

/** Where trackers.json, admins.json, name_map.json ... live (same as the job's settings). */
export const SETTINGS = Object.freeze({
  configSiteHost: 'mirageace.sharepoint.com',
  configSitePath: '/sites/Automation',
  configFolder: 'Application_Files/TrackerReminder',
});

const SITE_ADDRESS = window.location.origin + import.meta.env.BASE_URL;

export const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: SITE_ADDRESS,
    postLogoutRedirectUri: SITE_ADDRESS,
    navigateToLoginRequestUrl: false,
  },
  cache: { cacheLocation: 'localStorage' },
};

export const loginRequest = { scopes: SCOPES };
