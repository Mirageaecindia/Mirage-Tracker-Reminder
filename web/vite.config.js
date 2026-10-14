// ================================================================
// web/vite.config.js
// ----------------------------------------------------------------
// PURPOSE : Build / dev-server settings for the admin web page.
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * base = '/Mirage-Tracker-Reminder/': GitHub Pages serves the site
//       under the repo name (https://mirageaecindia-dev.github.io/
//       Mirage-Tracker-Reminder/). Must match the redirect URI IT registered
//       for the Admin app, or sign-in fails.
//     * Dev server on port 5173, strictPort: http://localhost:5173/ is the
//       second redirect URI registered - another port would break sign-in.
//     * base is '/' on the dev server and '/Mirage-Tracker-Reminder/' only for
//       the published build. WHY: the registered local redirect URI is
//       http://localhost:5173/ - with the repo-name base locally, sign-in
//       would return to a path the dev server does not serve.
//     * server.fs.allow '..': the page imports ../core (the SAME logic files
//       as the nightly job - no second copy of the rules).
// ================================================================

import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig(({ command }) => ({
  base: command === 'build' ? '/Mirage-Tracker-Reminder/' : '/',
  plugins: [react()],
  server: {
    port: 5173,
    strictPort: true,
    fs: { allow: ['..'] },
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
  },
}));
