// ================================================================
// htmlBuilder.js
// ----------------------------------------------------------------
// PURPOSE : Small helpers to write email HTML (tables, cells, encoding)
//           with the SAME look in every email (reminders + lead summary).
//           Only formatting (SRP); what goes in an email is the composers' job.
//           MIRRORS Excel_Notify\Services\HtmlBuilder.cs - keep both in step,
//           so reminder emails look exactly like Excel Notifier emails.
//
// USAGE   : every helper APPENDS to a string array (like C# StringBuilder);
//           the composer joins it once at the end.
//
// CHANGE LOG
// - 2026-10-14 : Initial version, ported 1:1 from HtmlBuilder.cs.
//     * Inline styles only. WHY: Outlook and many mail clients strip
//       <style> blocks; inline styles survive.
//     * Every value is HTML-encoded (encode). WHY: drawing / person names may
//       contain '<' or '&' and would break the table otherwise.
//     * date() takes "YYYY-MM-DD" strings (see dateCellParser.js) and prints
//       dd-MM-yyyy, the C# DateFormat.
//     * Same colours as C#: blue header #1e5aaa, overdue rows red #fdecea.
// ================================================================

export const DATE_FORMAT = 'dd-MM-yyyy';

const CELL = 'border:1px solid #c8c8c8;padding:4px 8px;';
const HEAD_CELL = CELL + 'background:#1e5aaa;color:#ffffff;text-align:left;';
const LABEL_CELL = CELL + 'background:#eef3fa;font-weight:bold;';
const HIGHLIGHT_CELL = 'background:#fdecea;color:#b00020;'; // overdue rows

/** HTML-encode; null/undefined become empty. */
export function encode(text) {
  return String(text ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Opening div with the email font. Close with endBody. */
export function beginBody(html) {
  html.push('<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;">');
}

export function endBody(html, footer) {
  html.push(`<p style="color:#666666;font-size:11px;">${encode(footer)}</p></div>`);
}

export function paragraph(html, text) {
  html.push(`<p>${encode(text)}</p>`);
}

export function heading(html, text) {
  html.push(`<h3 style="font-size:15px;margin:16px 0 6px 0;color:#1e5aaa;">${encode(text)}</h3>`);
}

/** Smaller heading (discipline under a role heading). */
export function subHeading(html, text) {
  html.push(`<h4 style="font-size:13px;margin:10px 0 4px 0;color:#333333;">${encode(text)}</h4>`);
}

/** Two-column "label | value" table (summary boxes). rows = [[label, value], ...]. */
export function labelValueTable(html, rows) {
  html.push('<table style="border-collapse:collapse;margin-bottom:12px;">');
  for (const [label, value] of rows) {
    html.push(`<tr><td style="${LABEL_CELL}">${encode(label)}</td><td style="${CELL}">${encode(value)}</td></tr>`);
  }
  html.push('</table>');
}

/** "YYYY-MM-DD" -> "dd-MM-yyyy", or the fallback text when there is no date. */
export function date(iso, fallback) {
  if (!iso) return fallback;
  const [y, m, d] = iso.split('-');
  return `${d}-${m}-${y}`;
}

/**
 * Table with a blue header row. Each row is one array of cell texts.
 * @param {(rowIndex: number) => boolean} [highlight]  true = row shown red (e.g. overdue)
 */
export function table(html, headers, rows, highlight) {
  html.push('<table style="border-collapse:collapse;margin-bottom:12px;"><tr>');
  for (const h of headers) html.push(`<th style="${HEAD_CELL}">${encode(h)}</th>`);
  html.push('</tr>');

  rows.forEach((row, index) => {
    const style = highlight && highlight(index) ? CELL + HIGHLIGHT_CELL : CELL;
    html.push('<tr>');
    for (const value of row) html.push(`<td style="${style}">${encode(value)}</td>`);
    html.push('</tr>');
  });
  html.push('</table>');
}
