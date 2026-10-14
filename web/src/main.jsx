// ================================================================
// web/src/main.jsx
// ----------------------------------------------------------------
// PURPOSE : Start the admin page: Microsoft sign-in (MSAL) around the app,
//           Fluent UI theme (Microsoft look, Mirage colours) around that.
//           Only wiring (SRP) - screens are in App.jsx / pages/.
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * msal.initialize() + handleRedirectPromise() BEFORE the first render.
//       WHY: sign-in uses the REDIRECT flow (works inside the Excel side
//       panel later, where pop-ups are blocked); the page must finish reading
//       the sign-in answer from the URL before React draws anything.
//     * The signed-in account is set as "active" so every token request
//       knows whose token it is.
// ================================================================

import React from 'react';
import { createRoot } from 'react-dom/client';
import { PublicClientApplication, EventType } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import { FluentProvider } from '@fluentui/react-components';

import { msalConfig } from './authConfig.js';
import { mirageTheme } from './theme.js';
import App from './App.jsx';

const msal = new PublicClientApplication(msalConfig);

async function start() {
  await msal.initialize();
  const result = await msal.handleRedirectPromise();
  if (result && result.account) msal.setActiveAccount(result.account);
  else if (!msal.getActiveAccount() && msal.getAllAccounts().length > 0) msal.setActiveAccount(msal.getAllAccounts()[0]);

  msal.addEventCallback((event) => {
    if (event.eventType === EventType.LOGIN_SUCCESS && event.payload && event.payload.account) {
      msal.setActiveAccount(event.payload.account);
    }
  });

  createRoot(document.getElementById('root')).render(
    <React.StrictMode>
      <MsalProvider instance={msal}>
        <FluentProvider theme={mirageTheme} style={{ minHeight: '100vh' }}>
          <App />
        </FluentProvider>
      </MsalProvider>
    </React.StrictMode>,
  );
}

start();
