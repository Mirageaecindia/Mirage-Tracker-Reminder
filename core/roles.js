// ================================================================
// roles.js
// ----------------------------------------------------------------
// PURPOSE : The four kinds of people a tracker assigns drawings to.
//           One list + display names, in ONE place (SRP).
//           MIRRORS Excel_Notify\Models\Role.cs - keep both in step.
//
// CHANGE LOG
// - 2026-10-13 : Initial version, ported 1:1 from Role.cs.
//     * Order of ALL_ROLES = order of the sections in every email and report
//       (same as the C# enum values 0..3).
//     * Which column belongs to which role is decided by roleMatcher.js
//       (keywords) or by the lead on the admin page - NOT here (SRP).
//     * QA/QC last, matching the tracker's column order.
//     * ADDED roleFromName(): saved setups store the role NAME
//       ("Engineers"), same as Excel_Notify's SheetSetupStore (Role.ToString()),
//       so a saved file is readable by a human and stable if values change.
// - 2026-10-13 : CommonJS (require / module.exports) -> ES module
//       (export). Logic unchanged.
//     * WHY: the same file must load in the browser (admin page, no build
//       step) and in Node - see package.json "type": "module".
// ================================================================

/** Role ids. Values match the C# enum so saved data stays comparable. */
export const Role = Object.freeze({
  Engineers: 0,
  Modelers: 1,
  Annotation: 2,
  QaQc: 3,
});

/** All roles in display order. */
export const ALL_ROLES = Object.freeze([Role.Engineers, Role.Modelers, Role.Annotation, Role.QaQc]);

/** Short name for checkboxes, grid headers and emails. */
export function roleShortName(role) {
  switch (role) {
    case Role.Engineers: return 'Engineers';
    case Role.Modelers: return 'Modelers';
    case Role.QaQc: return 'QA/QC';
    default: return 'Annotation'; // same fallback as RoleNames.Short in C#
  }
}

/** Enum name as saved in files ("QaQc"), like C# Role.ToString(). */
export function roleKey(role) {
  return Object.keys(Role).find((k) => Role[k] === role);
}

/** Saved enum name -> role id; undefined for an unknown name (ignored, like Enum.TryParse). */
export function roleFromName(name) {
  return Object.prototype.hasOwnProperty.call(Role, name) ? Role[name] : undefined;
}
