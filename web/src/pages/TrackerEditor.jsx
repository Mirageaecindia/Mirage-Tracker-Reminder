// ================================================================
// web/src/pages/TrackerEditor.jsx
// ----------------------------------------------------------------
// PURPOSE : Add or edit ONE tracker:
//             1. paste the SharePoint link -> read the workbook
//             2. every sheet's layout is auto-detected with the SAME rules as
//                Excel_Notify (core/sheetLayoutDetector + roleMatcher); the
//                admin confirms or changes which column block each of the
//                4 roles uses, per sheet, or leaves a sheet out
//             3. name, lead, send time (IST), test mode, sender mode,
//                overdue-backlog choice
//             4. Save (validated by core/trackerConfig - the job's own check)
//           Only UI + building the tracker object (SRP); App.jsx saves it.
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Columns are stored by block TITLE (Excel_Notify's SheetSetup rule):
//       an inserted column does not break the setup.
//     * DEFAULT for a new tracker: every auto-matched role ticked.
//       DIFFERENCE from Excel_Notify (which ticked only Engineers + Modelers
//       the first time): for reminders an empty Name column simply produces
//       nothing, so reading all matched roles costs nothing and misses no one.
//     * When editing, a saved column title that no longer exists in the
//       workbook is shown in red - the job would skip that role (warning).
//     * registeredOn = today (India) for a new tracker; kept when editing
//       (the overdue-backlog rule depends on it).
//     * Delete needs a second click ("Click again to delete").
// ================================================================

import { useEffect, useMemo, useState } from 'react';
import {
  Button, Checkbox, Dropdown, Field, Input, MessageBar, MessageBarBody, Option, Radio,
  RadioGroup, Spinner, Switch, Table, TableBody, TableCell, TableHeader, TableHeaderCell,
  TableRow, Text, Title3,
} from '@fluentui/react-components';

import PersonPicker from '../components/PersonPicker.jsx';
import { resolveTrackerLink } from '../services/trackerLinks.js';
import { loadWorkbookLayouts } from '../services/trackerReading.js';
import { ALL_ROLES, roleKey, roleShortName } from '../../../core/roles.js';
import { validateTracker, withDefaults } from '../../../core/trackerConfig.js';
import { BacklogPolicy } from '../../../core/recipientBuilder.js';
import { SenderMode } from '../../../core/emailComposer.js';
import { DEFAULT_SEND_TIME, indiaNow, sendTimeLabel } from '../../../core/schedule.js';

const NOT_USED = '';
const SEND_TIMES = Array.from({ length: 48 }, (_, i) => `${String(Math.floor(i / 2)).padStart(2, '0')}:${i % 2 ? '30' : '00'}`);

const BACKLOG_TEXT = {
  [BacklogPolicy.AfterRegistration]: 'Do not email drawings that were ALREADY overdue when this tracker was registered (recommended)',
  [BacklogPolicy.EmailOnce]: 'Email those old overdue drawings once, like any other overdue drawing',
  [BacklogPolicy.LeadSummaryOnly]: 'List those old overdue drawings in the lead\'s summary only - do not email people',
};

/**
 * @param {{ graph: object, directory: object[], me: {name: string, email: string},
 *           trackers: object[], editing: object|null, busy: boolean,
 *           onSave: (t: object) => void, onCancel: () => void, onDelete: (id: string) => void }} props
 */
export default function TrackerEditor({ graph, directory, me, trackers, editing, busy, onSave, onCancel, onDelete }) {
  const start = editing ? withDefaults(editing) : null;

  const [link, setLink] = useState('');
  const [loc, setLoc] = useState(start ? { siteHost: start.siteHost, sitePath: start.sitePath, filePath: start.filePath, fileName: start.filePath.split('/').pop() } : null);
  const [layouts, setLayouts] = useState(null); // Map sheetName -> SheetLayout
  const [choices, setChoices] = useState(start ? choicesFromTracker(start) : {});
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const [name, setName] = useState(start ? start.name : '');
  const [lead, setLead] = useState(start ? start.lead : me);
  const [sendTime, setSendTime] = useState(start ? start.sendTime : DEFAULT_SEND_TIME);
  const [testMode, setTestMode] = useState(start ? start.testMode : true);
  const [testRecipient, setTestRecipient] = useState(start && start.testRecipient ? personOf(directory, start.testRecipient) : me);
  const [senderMode, setSenderMode] = useState(start ? start.senderMode : SenderMode.Default);
  const [onBehalfOf, setOnBehalfOf] = useState(start ? start.onBehalfOf : null);
  const [backlogPolicy, setBacklogPolicy] = useState(start ? start.backlogPolicy : BacklogPolicy.AfterRegistration);
  const [confirmDelete, setConfirmDelete] = useState(false);

  // Editing: read the workbook straight away so the column lists are available.
  useEffect(() => {
    if (start) readWorkbook(loc, false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function readWorkbook(location, applyAutoMatch) {
    setLoading(true);
    setError(null);
    try {
      const found = await loadWorkbookLayouts(graph, location);
      setLayouts(found);
      if (applyAutoMatch) setChoices(choicesFromAutoMatch(found));
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  async function onReadLink() {
    setError(null);
    try {
      const resolved = await resolveTrackerLink(graph, link);
      setLoc(resolved);
      if (!name) setName(resolved.fileName.replace(/\.xls[xm]$/i, ''));
      await readWorkbook(resolved, true);
    } catch (e) {
      setError(e.message);
    }
  }

  const tracker = useMemo(() => {
    if (!loc) return null;
    const sheets = Object.entries(choices)
      .filter(([, c]) => c.include)
      .map(([sheetName, c]) => ({
        name: sheetName,
        roles: Object.fromEntries(Object.entries(c.roles).filter(([, title]) => title)),
      }))
      .filter((s) => Object.keys(s.roles).length > 0);
    return {
      id: start ? start.id : uniqueId(name, trackers),
      name: name.trim(),
      enabled: start ? start.enabled : true,
      testMode,
      testRecipient: testRecipient ? testRecipient.email : '',
      siteHost: loc.siteHost,
      sitePath: loc.sitePath,
      filePath: loc.filePath,
      lead,
      senderMode,
      onBehalfOf: senderMode === SenderMode.OnBehalf ? onBehalfOf : null,
      backlogPolicy,
      registeredOn: start ? start.registeredOn : indiaNow().date,
      sendTime,
      sheets,
    };
  }, [loc, choices, name, lead, sendTime, testMode, testRecipient, senderMode, onBehalfOf, backlogPolicy, start, trackers]);

  const problems = tracker ? validateTracker(withDefaults(tracker)) : [];

  return (
    <div style={{ display: 'grid', gap: 16, maxWidth: 1100 }}>
      <Title3>{start ? `Edit tracker - ${start.name}` : 'Add tracker'}</Title3>
      {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}

      {/* ---- 1. the file ---- */}
      <Field label="1. Tracker file - paste its SharePoint link (address bar or 'Copy link')">
        <div style={{ display: 'flex', gap: 8 }}>
          <Input style={{ flex: 1 }} value={link} onChange={(_, d) => setLink(d.value)}
            placeholder={loc ? `${loc.sitePath}/${loc.filePath}` : 'https://mirageace.sharepoint.com/...'} />
          <Button onClick={onReadLink} disabled={!link.trim() || loading}>{loc ? 'Use this file instead' : 'Read workbook'}</Button>
          {loc && <Button onClick={() => readWorkbook(loc, false)} disabled={loading}>Re-read columns</Button>}
        </div>
      </Field>
      {loc && <Text size={200}>File: {loc.sitePath} / {loc.filePath}</Text>}
      {loading && <Spinner label="Reading the workbook..." />}

      {/* ---- 2. sheets + columns ---- */}
      {layouts && (
        <Field label="2. Sheets and columns - which column block holds each role's Name / Date / Status">
          <Table size="small">
            <TableHeader>
              <TableRow>
                <TableHeaderCell style={{ width: 200 }}>Sheet</TableHeaderCell>
                {ALL_ROLES.map((r) => <TableHeaderCell key={r}>{roleShortName(r)}</TableHeaderCell>)}
              </TableRow>
            </TableHeader>
            <TableBody>
              {[...layouts.values()].map((layout) => (
                <SheetRow key={layout.sheetName} layout={layout}
                  choice={choices[layout.sheetName]}
                  onChange={(c) => setChoices({ ...choices, [layout.sheetName]: c })} />
              ))}
            </TableBody>
          </Table>
        </Field>
      )}

      {/* ---- 3. settings ---- */}
      {loc && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
          <Field label="Tracker name (shown in emails)">
            <Input value={name} onChange={(_, d) => setName(d.value)} />
          </Field>
          <Field label="Lead (gets the nightly summary; replies go to them)">
            <PersonPicker directory={directory} value={lead} onChange={setLead} />
          </Field>
          <Field label="Send time (India time)">
            <Dropdown value={sendTimeLabel(sendTime)} selectedOptions={[sendTime]}
              onOptionSelect={(_, d) => setSendTime(d.optionValue)}>
              {SEND_TIMES.map((t) => <Option key={t} value={t}>{sendTimeLabel(t)}</Option>)}
            </Dropdown>
          </Field>
          <Field label="Test mode (every email goes ONLY to the test person)">
            <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
              <Switch checked={testMode} onChange={(_, d) => setTestMode(d.checked)} label={testMode ? 'On' : 'Off - LIVE'} />
              {testMode && <PersonPicker directory={directory} value={testRecipient} onChange={setTestRecipient} />}
            </div>
          </Field>
          <Field label="Sender">
            <RadioGroup value={senderMode} onChange={(_, d) => setSenderMode(d.value)}>
              <Radio value={SenderMode.Default} label="Mirage AEC - Tracker Reminders (replies go to the lead)" />
              <Radio value={SenderMode.OnBehalf} label="On behalf of a person (their name in the email, replies go to them)" />
            </RadioGroup>
            {senderMode === SenderMode.OnBehalf && <PersonPicker directory={directory} value={onBehalfOf} onChange={setOnBehalfOf} />}
          </Field>
          <Field label="Drawings already overdue when the tracker is registered">
            <RadioGroup value={backlogPolicy} onChange={(_, d) => setBacklogPolicy(d.value)}>
              {Object.values(BacklogPolicy).map((p) => <Radio key={p} value={p} label={BACKLOG_TEXT[p]} />)}
            </RadioGroup>
          </Field>
        </div>
      )}

      {tracker && problems.length > 0 && (
        <MessageBar intent="warning"><MessageBarBody>
          Fix before saving:<ul style={{ margin: '4px 0 0 0' }}>{problems.map((p) => <li key={p}>{p}</li>)}</ul>
        </MessageBarBody></MessageBar>
      )}

      <div style={{ display: 'flex', gap: 8 }}>
        <Button appearance="primary" disabled={!tracker || problems.length > 0 || busy || loading} onClick={() => onSave(tracker)}>Save</Button>
        <Button onClick={onCancel} disabled={busy}>Cancel</Button>
        <span style={{ flex: 1 }} />
        {start && (
          <Button disabled={busy} onClick={() => (confirmDelete ? onDelete(start.id) : setConfirmDelete(true))}>
            {confirmDelete ? 'Click again to delete this tracker' : 'Delete tracker'}
          </Button>
        )}
      </div>
    </div>
  );
}

/** One sheet: include tick-box + a column-block dropdown per role. */
function SheetRow({ layout, choice, onChange }) {
  if (!layout.isTracker) {
    return (
      <TableRow>
        <TableCell><Text style={{ color: '#888' }}>{layout.sheetName}</Text></TableCell>
        <TableCell colSpan={4}><Text style={{ color: '#888' }}>not a tracker ({layout.notTrackerReason})</Text></TableCell>
      </TableRow>
    );
  }
  const c = choice || { include: false, roles: {} };
  return (
    <TableRow>
      <TableCell>
        <Checkbox checked={c.include} label={layout.sheetName} onChange={(_, d) => onChange({ ...c, include: Boolean(d.checked) })} />
      </TableCell>
      {ALL_ROLES.map((role) => {
        const key = roleKey(role);
        const title = c.roles[key] || NOT_USED;
        const missing = title && !layout.findBlock(title);
        return (
          <TableCell key={key}>
            <Dropdown size="small" disabled={!c.include}
              value={title ? (missing ? `"${title}" - NOT FOUND` : String(layout.findBlock(title))) : '(not used)'}
              selectedOptions={[title]}
              style={missing ? { color: '#b00020', minWidth: 160 } : { minWidth: 160 }}
              onOptionSelect={(_, d) => onChange({ ...c, roles: { ...c.roles, [key]: d.optionValue } })}>
              <Option value={NOT_USED}>(not used)</Option>
              {layout.blocks.map((b) => <Option key={b.title} value={b.title}>{String(b)}</Option>)}
            </Dropdown>
          </TableCell>
        );
      })}
    </TableRow>
  );
}

/** New tracker: every tracker sheet ticked, each role = its keyword match (see change log). */
function choicesFromAutoMatch(layouts) {
  const out = {};
  for (const layout of layouts.values()) {
    if (!layout.isTracker) continue;
    const roles = {};
    for (const [role, block] of layout.autoMatched) roles[roleKey(role)] = block.title;
    out[layout.sheetName] = { include: true, roles };
  }
  return out;
}

function choicesFromTracker(t) {
  const out = {};
  for (const s of t.sheets) out[s.name] = { include: true, roles: { ...s.roles } };
  return out;
}

function uniqueId(name, trackers) {
  const base = (name || 'tracker').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 30) || 'tracker';
  let id = base.length >= 2 ? base : `${base}-1`;
  for (let i = 2; trackers.some((t) => t.id === id); i++) id = `${base}-${i}`;
  return id;
}

function personOf(directory, email) {
  const u = directory.find((x) => x.email.toLowerCase() === String(email).toLowerCase());
  return u ? { name: u.displayName, email: u.email } : { name: email, email };
}
