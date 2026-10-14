// ================================================================
// web/src/theme.js
// ----------------------------------------------------------------
// PURPOSE : Fluent UI (Microsoft) look with Mirage colours (user decision
//           2026-10-14): navy brand colour for buttons / links / focus, gold
//           accent for the header line. Only styling (SRP).
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Fluent builds every brand shade from a 16-step ramp (10 = darkest,
//       160 = lightest); 80 is the main button colour. Ramp built around
//       Mirage navy (#1E3D75 at 60, buttons at 80 for readable contrast).
//     * GOLD / NAVY exported for the header bar (App.jsx).
// ================================================================

import { createLightTheme } from '@fluentui/react-components';

export const NAVY = '#0F1F3C';
export const GOLD = '#C9A24A';

const mirageBrand = {
  10: '#050B16', 20: '#0A1529', 30: '#0F1F3C', 40: '#14294F',
  50: '#193362', 60: '#1E3D75', 70: '#234788', 80: '#28519B',
  90: '#3D63A8', 100: '#5276B5', 110: '#6789C2', 120: '#7C9CCF',
  130: '#91AFDB', 140: '#A6C2E8', 150: '#BBD5F4', 160: '#D0E8FF',
};

export const mirageTheme = createLightTheme(mirageBrand);
