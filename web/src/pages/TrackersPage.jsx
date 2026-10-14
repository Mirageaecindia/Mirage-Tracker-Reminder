// ================================================================
// web/src/pages/TrackersPage.jsx
// ----------------------------------------------------------------
// PURPOSE : List of registered trackers with their key settings, plus
//           Add / Edit / Enable-Disable. Only UI (SRP): saving is done by
//           App.jsx (one place that owns trackers.json and its eTag).
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Problems in trackers.json (from core parseTrackersFile - the SAME
//       check the job runs) are shown at the top, so an admin sees what the
//       job will skip tonight before it happens.
//     * "Last ran" comes from runstate.json (written by the job).
// ================================================================

import {
  Badge, Button, MessageBar, MessageBarBody, Switch, Table, TableBody, TableCell,
  TableHeader, TableHeaderCell, TableRow, Text, Title3,
} from '@fluentui/react-components';
import { parseTrackersFile, withDefaults } from '../../../core/trackerConfig.js';
import { sendTimeLabel } from '../../../core/schedule.js';

/**
 * @param {{ trackersData: object|null, runState: object|null, busy: boolean,
 *           onAdd: () => void, onEdit: (id: string) => void,
 *           onToggleEnabled: (id: string, enabled: boolean) => void }} props
 */
export default function TrackersPage({ trackersData, runState, busy, onAdd, onEdit, onToggleEnabled }) {
  const raw = (trackersData && Array.isArray(trackersData.trackers) ? trackersData.trackers : []).map(withDefaults);
  const { problems } = parseTrackersFile(trackersData || { trackers: [] });
  const lastRun = (id) => (runState && runState.trackers && runState.trackers[id] ? runState.trackers[id].lastRunOn : null);

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12 }}>
        <Title3>Trackers</Title3>
        <Button appearance="primary" onClick={onAdd} disabled={busy}>Add tracker</Button>
      </div>

      {trackersData && problems.length > 0 && (
        <MessageBar intent="warning" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            The nightly job will SKIP these until fixed:
            <ul style={{ margin: '4px 0 0 0' }}>{problems.map((p) => <li key={p}>{p}</li>)}</ul>
          </MessageBarBody>
        </MessageBar>
      )}

      {raw.length === 0 ? (
        <Text>No trackers registered yet. Click "Add tracker" and paste a tracker's SharePoint link.</Text>
      ) : (
        <Table aria-label="Trackers">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Name</TableHeaderCell>
              <TableHeaderCell>File</TableHeaderCell>
              <TableHeaderCell>Sends at (IST)</TableHeaderCell>
              <TableHeaderCell>Mode</TableHeaderCell>
              <TableHeaderCell>Lead</TableHeaderCell>
              <TableHeaderCell>Last ran</TableHeaderCell>
              <TableHeaderCell>Enabled</TableHeaderCell>
              <TableHeaderCell />
            </TableRow>
          </TableHeader>
          <TableBody>
            {raw.map((t) => (
              <TableRow key={t.id}>
                <TableCell>{t.name}</TableCell>
                <TableCell title={`${t.sitePath}/${t.filePath}`}>{String(t.filePath || '').split('/').pop()}</TableCell>
                <TableCell>{sendTimeLabel(t.sendTime)}</TableCell>
                <TableCell>
                  {t.testMode
                    ? <Badge appearance="tint" color="warning">Test - all mail to {t.testRecipient}</Badge>
                    : <Badge appearance="tint" color="success">Live</Badge>}
                </TableCell>
                <TableCell>{t.lead ? t.lead.name : '-'}</TableCell>
                <TableCell>{lastRun(t.id) || 'never'}</TableCell>
                <TableCell>
                  <Switch checked={t.enabled} disabled={busy} onChange={(_, d) => onToggleEnabled(t.id, d.checked)} />
                </TableCell>
                <TableCell><Button onClick={() => onEdit(t.id)} disabled={busy}>Edit</Button></TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      )}
    </div>
  );
}
