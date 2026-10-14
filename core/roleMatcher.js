// ================================================================
// roleMatcher.js
// ----------------------------------------------------------------
// PURPOSE : Decide from a stage TITLE which role it belongs to, using the
//           user's keywords. Only matching (SRP). The lead can always
//           override on the admin page.
//             QA/QC      - "qa" / "qc" (whole words) or "quality"
//             Engineers  - "engineer"  (engineer, engineers, engineering)
//             Modelers   - "model"     (model, modeling, modeller, modelors)
//             Annotation - "annot"     (annotation)
//           MIRRORS Excel_Notify\Workbook\RoleMatcher.cs - keep both in step.
//
// RULES (same order as C#, order matters):
//   1. QA/QC checked FIRST, so "Enegineering QA/QC" -> QA/QC, and still
//      QA/QC once someone fixes the spelling to "Engineering QA/QC".
//   2. EXCLUDE words (check, review, print, remark, comment) -> never a role.
//      "Naviswork Checking" / "Print For 100% SD" are not roles.
//   3. engineer -> model -> annot.
//   "Inputs" alone (SE01 / SW01) is deliberately NOT a keyword (user
//   decision): the lead picks that column on the admin page.
//
// CHANGE LOG
// - 2026-10-13 : Initial version, ported 1:1 from RoleMatcher.cs.
//     * C# Regex IgnoreCase -> JS /i flag; \b works the same for these
//       ASCII words.
//     * Returns null for "no role" (C# returned Role? null).
// - 2026-10-13 : CommonJS (require / module.exports) -> ES module
//       (import / export). Logic unchanged.
//     * WHY: the same file must load in the browser and in Node - see
//       package.json "type": "module".
// ================================================================

import { Role } from './roles.js';

// QA/QC stage: "QA", "QC" as whole words, or "quality".
const QA_QC = /\bqa\b|\bqc\b|quality/i;

// Never a role (checked AFTER QA/QC - see rules).
const EXCLUDE = /check|review|print|remark|comment/i;

/**
 * The role a stage title belongs to, or null.
 * @param {string} title  header text, e.g. "Modeling with Coordination"
 * @returns {number|null} a Role value or null
 */
export function matchRole(title) {
  if (title == null || String(title).trim() === '') return null;
  const text = String(title);

  if (QA_QC.test(text)) return Role.QaQc; // FIRST (rule 1)
  if (EXCLUDE.test(text)) return null;    // rule 2

  const lower = text.toLowerCase();
  if (lower.includes('engineer')) return Role.Engineers;
  if (lower.includes('model')) return Role.Modelers;
  if (lower.includes('annot')) return Role.Annotation;
  return null;
}
