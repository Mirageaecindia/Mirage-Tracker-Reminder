// ================================================================
// job/mailSender.js
// ----------------------------------------------------------------
// PURPOSE : Send ONE composed email from the shared mailbox
//           (project-reminders@) through Microsoft Graph:
//             POST /users/{sender}/sendMail
//           Only sends (SRP). Who gets what = recipientBuilder; the text =
//           the composers.
//           Equivalent of Excel_Notify\Services\GraphMailSender.cs (same
//           checks), but app-only from a shared mailbox instead of /me.
//
// CHANGE LOG
// - 2026-10-14 : Initial version, adapted from GraphMailSender.cs.
//     * Domain check (@mirageaec.com) on EVERY To and Reply-To, here in the
//       last step, so no caller can bypass it. One bad address blocks the
//       whole email (never send half-checked) - same rule as C#.
//     * Size check BEFORE sending (limit 3.5 MB, Graph rejects ~4 MB) with a
//       clear message instead of Graph's generic 413 - same as C#.
//     * Logo: inline attachment, contentBytes as base64 STRING (what Graph
//       expects), isInline + contentId = the cid the HTML uses. Missing logo
//       file -> the <img> is removed and the email still goes (C# rule:
//       a missing picture must never stop the reminders).
//     * TEST MODE: To is replaced by the test address and the subject says
//       who it WOULD have gone to: "[TEST -> abhi@..] Reminder: ...".
//       Reply-To is kept, so a test reply shows where real replies would go.
//     * DRY RUN: prints what would be sent, sends nothing.
//     * The shared mailbox keeps a copy in its Sent Items (saveToSentItems).
//     * Permission: the job app may ONLY send as project-reminders@ (Exchange
//       RBAC scope, verified 2026-10-13). Any other sender -> Graph 403.
// ================================================================

import { isAllowedAddress } from '../core/recipientResolver.js';

const MAX_REQUEST_BYTES = 3500000;

export class MailSender {
  /**
   * @param {import('./graphClient.js').GraphClient} graph
   * @param {{senderMailbox: string, logoBytes: Buffer|null, dryRun: boolean, log?: (s: string) => void}} opts
   */
  constructor(graph, opts) {
    this.graph = graph;
    this.opts = opts;
    this.log = opts.log || console.log;
  }

  /**
   * @param {{to: object[], replyTo: object[], subject: string, html: string, inlineImages: object[]}} mail
   * @param {string|null} testRecipient  non-null = TEST MODE (see change log)
   */
  async send(mail, testRecipient) {
    let { to, subject, html } = mail;
    if (testRecipient) {
      subject = `[TEST -> ${to.map((t) => t.email).join(', ')}] ${subject}`;
      to = [{ name: 'Tracker Reminder test', email: testRecipient }];
    }

    if (to.length === 0) throw new Error('Email has no recipient.');
    for (const r of [...to, ...mail.replyTo]) {
      if (!isAllowedAddress(r.email)) throw new Error(`Blocked: "${r.email}" is not a @mirageaec.com address.`);
    }

    const attachments = [];
    for (const img of mail.inlineImages || []) {
      if (!this.opts.logoBytes) {
        html = html.replace(new RegExp(`<p><img src="cid:${img.contentId}"[^>]*></p>`, 'g'), ''); // no logo file
        continue;
      }
      attachments.push({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: img.fileName,
        contentType: img.contentType,
        contentBytes: this.opts.logoBytes.toString('base64'),
        isInline: true,
        contentId: img.contentId,
      });
    }

    const payload = {
      message: {
        subject,
        body: { contentType: 'HTML', content: html },
        toRecipients: to.map(asRecipient),
        ...(mail.replyTo.length ? { replyTo: mail.replyTo.map(asRecipient) } : {}),
        ...(attachments.length ? { attachments } : {}),
      },
      saveToSentItems: true,
    };

    const bytes = Buffer.byteLength(JSON.stringify(payload), 'utf8');
    if (bytes > MAX_REQUEST_BYTES) {
      throw new Error(`Email to ${to.map((t) => t.email).join(', ')} is too large (${Math.round(bytes / 1024)} KB, limit ~3,400 KB).`);
    }

    if (this.opts.dryRun) {
      this.log(`   [dry run] would send: ${subject}  ->  ${to.map((t) => t.email).join(', ')}`);
      return;
    }
    await this.graph.postJson(`users/${encodeURIComponent(this.opts.senderMailbox)}/sendMail`, payload);
  }
}

function asRecipient(r) {
  return { emailAddress: { address: r.email.trim(), ...(r.name ? { name: r.name } : {}) } };
}
