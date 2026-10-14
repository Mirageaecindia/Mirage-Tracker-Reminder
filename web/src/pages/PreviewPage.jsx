// ================================================================
// web/src/pages/PreviewPage.jsx
// ----------------------------------------------------------------
// PURPOSE : Show EXACTLY the emails the job would send for one tracker on a
//           chosen date - same core code as job/runNightly.js (read ->
//           resolve -> decide -> compose), with the real sent log, so
//           "already reminded" items are left out just like at night.
//           NOTHING is sent and nothing is saved.
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * The page cannot see GitHub's GLOBAL_TEST_MODE setting, so it uses the
//       tracker's own testMode to pick the sent-log namespace ("<id>#test"
//       in test mode) and to show the "[TEST -> ...]" subject. If
//       GLOBAL_TEST_MODE is on, the real run is in test mode regardless.
//     * Emails are shown in a sandboxed iframe (no scripts) with the logo
//       swapped from cid: to the page's own logo file - the Excel_Notify
//       PreviewHtml trick.
// ================================================================

import { useState } from 'react';
import {
  Button, Dropdown, Field, Input, MessageBar, MessageBarBody, Option, Spinner, Table,
  TableBody, TableCell, TableHeader, TableHeaderCell, TableRow, Text, Title3,
} from '@fluentui/react-components';

import { readTrackerForPage } from '../services/trackerReading.js';
import { RecipientResolver, parseNameMap } from '../../../core/recipientResolver.js';
import { buildReminders } from '../../../core/recipientBuilder.js';
import { SentLog } from '../../../core/sentLog.js';
import { composeReminder, LOGO } from '../../../core/emailComposer.js';
import { composeLeadSummary, shouldSendLeadSummary } from '../../../core/leadSummaryComposer.js';
import { withDefaults } from '../../../core/trackerConfig.js';
import { indiaNow } from '../../../core/schedule.js';
import logoUrl from '../../../assets/mirage_logo.jpg';

/**
 * @param {{ graph: object, directory: object[], trackers: object[],
 *           nameMapData: object[]|null, sentLogData: object|null }} props
 */
export default function PreviewPage({ graph, directory, trackers, nameMapData, sentLogData }) {
  const [trackerId, setTrackerId] = useState(trackers[0] ? trackers[0].id : '');
  const [date, setDate] = useState(indiaNow().date);
  const [mails, setMails] = useState(null);
  const [plan, setPlan] = useState(null);
  const [shown, setShown] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  async function build() {
    setLoading(true);
    setError(null);
    setMails(null);
    try {
      const t = withDefaults(trackers.find((x) => x.id === trackerId));
      const { read, people, workbookName } = await readTrackerForPage(graph, t);
      const resolver = new RecipientResolver(directory, [parseNameMap(nameMapData || []).map, people.map]);
      const result = buildReminders({
        items: read.items, today: date, trackerId: t.testMode ? `${t.id}#test` : t.id,
        registeredOn: t.registeredOn, backlogPolicy: t.backlogPolicy, resolver, sentLog: new SentLog(sentLogData),
      });
      const list = result.recipients.map((recipient) => {
        const mail = composeReminder(recipient, {
          workbookName, today: date, senderMode: t.senderMode, lead: t.lead, onBehalfOf: t.onBehalfOf, includeLogo: true,
        });
        return testView(mail, t);
      });
      const summaryInput = {
        trackerName: t.name, workbookName, today: date, lead: t.lead,
        results: result.recipients.map((recipient) => ({ recipient, ok: true })),
        unmatched: result.unmatched, backlog: result.backlog, alreadySentCount: result.alreadySentCount,
        warnings: [...read.warnings, ...people.warnings], errors: [], testMode: t.testMode, includeLogo: true,
      };
      if (shouldSendLeadSummary(summaryInput)) list.push(testView(composeLeadSummary(summaryInput), t));
      setPlan(result);
      setMails(list);
      setShown(0);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  const current = mails && mails[shown];

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <Title3>Preview</Title3>
      <Text>Shows what the job would send. Nothing is sent from this screen.</Text>
      {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}

      <div style={{ display: 'flex', gap: 12, alignItems: 'end' }}>
        <Field label="Tracker">
          <Dropdown value={(trackers.find((t) => t.id === trackerId) || {}).name || ''} selectedOptions={[trackerId]}
            onOptionSelect={(_, d) => setTrackerId(d.optionValue)}>
            {trackers.map((t) => <Option key={t.id} value={t.id}>{t.name}</Option>)}
          </Dropdown>
        </Field>
        <Field label="Date (India)">
          <Input type="date" value={date} onChange={(_, d) => setDate(d.value)} />
        </Field>
        <Button appearance="primary" onClick={build} disabled={!trackerId || !date || loading}>Build preview</Button>
      </div>

      {loading && <Spinner label="Reading the tracker..." />}

      {plan && (
        <Text>
          {plan.recipients.length} reminder email(s), {plan.unmatched.length} name(s) not matched,
          {' '}{plan.alreadySentCount} item(s) already reminded (left out), {plan.backlog.length} old overdue item(s) for the lead only.
        </Text>
      )}

      {mails && mails.length === 0 && <Text>Nothing would be sent for this tracker on that date.</Text>}

      {mails && mails.length > 0 && (
        <div style={{ display: 'grid', gridTemplateColumns: '420px 1fr', gap: 16 }}>
          <Table size="small">
            <TableHeader>
              <TableRow><TableHeaderCell>To</TableHeaderCell><TableHeaderCell>Subject</TableHeaderCell></TableRow>
            </TableHeader>
            <TableBody>
              {mails.map((m, i) => (
                <TableRow key={i} onClick={() => setShown(i)} style={{ cursor: 'pointer', background: i === shown ? '#eef3fa' : undefined }}>
                  <TableCell>{m.toText}</TableCell>
                  <TableCell>{m.subject}</TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
          {current && (
            <div>
              <Text size={200}>To: {current.toText} | Reply-To: {current.replyText || '-'}</Text>
              <iframe title="Email preview" sandbox="" srcDoc={current.html}
                style={{ width: '100%', height: 640, border: '1px solid #ddd', background: '#fff', marginTop: 6 }} />
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/** Composed mail -> what the inbox would show (test redirect + logo for the iframe). */
function testView(mail, tracker) {
  const realTo = mail.to.map((t) => t.email).join(', ');
  return {
    toText: tracker.testMode ? `${tracker.testRecipient} (test; real: ${realTo})` : realTo,
    replyText: mail.replyTo.map((r) => r.email).join(', '),
    subject: tracker.testMode ? `[TEST -> ${realTo}] ${mail.subject}` : mail.subject,
    html: mail.html.split(`cid:${LOGO.contentId}`).join(new URL(logoUrl, window.location.href).href),
  };
}

