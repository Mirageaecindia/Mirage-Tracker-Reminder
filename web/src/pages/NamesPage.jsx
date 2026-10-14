// ================================================================
// web/src/pages/NamesPage.jsx
// ----------------------------------------------------------------
// PURPOSE : Name matching - the page version of Excel_Notify's "Pick..."
//           button + name-map Export / Import / Delete:
//             - pick a tracker -> every name in it, with HOW it matches today
//               (name map / People sheet / company directory / NOT matched)
//             - map any name to a person (saved to the SHARED name_map.json)
//             - remove a mapping, import an Excel_Notify name map file
//           Matching itself = core/recipientResolver (the job's own rules).
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Changes are collected and saved with ONE "Save" (one eTag check),
//       not one save per click.
//     * Import MERGES like Excel_Notify NameMapStore.Import: new name added,
//       same name with a different email UPDATED (imported wins), others
//       kept; invalid / non-@mirageaec.com entries skipped and counted.
//     * A mapped email that is not in the directory is shown as
//       "person not found" - the job would fall through to the next source.
// ================================================================

import { useMemo, useState } from 'react';
import {
  Badge, Button, Dropdown, MessageBar, MessageBarBody, Option, Spinner, Table, TableBody,
  TableCell, TableHeader, TableHeaderCell, TableRow, Text, Title3,
} from '@fluentui/react-components';

import PersonPicker from '../components/PersonPicker.jsx';
import { readTrackerForPage } from '../services/trackerReading.js';
import { RecipientResolver, parseNameMap, nameMapKey, nameMapToEntries } from '../../../core/recipientResolver.js';

/**
 * @param {{ graph: object, directory: object[], trackers: object[], nameMapData: object[]|null,
 *           busy: boolean, onSave: (entries: object[], message: string) => Promise<boolean> }} props
 */
export default function NamesPage({ graph, directory, trackers, nameMapData, busy, onSave }) {
  const saved = useMemo(() => parseNameMap(nameMapData || []).map, [nameMapData]);
  const [draft, setDraft] = useState(null); // Map being edited (null = no unsaved changes)
  const map = draft || saved;

  const [trackerId, setTrackerId] = useState(trackers[0] ? trackers[0].id : '');
  const [names, setNames] = useState(null); // [{ name, count, peopleEmail }]
  const [peopleMap, setPeopleMap] = useState(new Map());
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [info, setInfo] = useState(null);

  async function loadNames() {
    setLoading(true);
    setError(null);
    try {
      const t = trackers.find((x) => x.id === trackerId);
      const { read, people } = await readTrackerForPage(graph, t);
      const counts = new Map();
      for (const item of read.items) {
        const n = item.personName.trim();
        counts.set(n, (counts.get(n) || 0) + 1);
      }
      setPeopleMap(people.map);
      setNames([...counts.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([name, count]) => ({ name, count })));
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  /** How a name matches today: { source, user } (see change log). */
  function matchOf(name) {
    const mapped = map.get(nameMapKey(name));
    if (mapped) {
      const user = directory.find((u) => u.email.toLowerCase() === mapped.toLowerCase());
      return { source: user ? 'Name map' : 'Name map - person not found', user, mapped };
    }
    const fromPeople = new RecipientResolver(directory, [peopleMap]).resolve(name);
    if (fromPeople && peopleMap.has(nameMapKey(name))) return { source: 'People sheet', user: fromPeople };
    const fromDirectory = new RecipientResolver(directory, []).resolve(name);
    if (fromDirectory) return { source: 'Directory', user: fromDirectory };
    return { source: 'NOT matched', user: null };
  }

  function setMapping(name, person) {
    const next = new Map(map);
    if (person) next.set(nameMapKey(name), person.email);
    else next.delete(nameMapKey(name));
    setDraft(next);
  }

  async function importFile(file) {
    setError(null);
    try {
      if (!/\.json$/i.test(file.name) || file.size > 1024 * 1024) throw new Error('Choose a .json name map file (max 1 MB).');
      const { map: incoming, rejected } = parseNameMap(JSON.parse(await file.text()));
      if (incoming.size === 0) throw new Error('No valid entries in that file (each needs Name + an @mirageaec.com Email).');
      const next = new Map(map);
      let added = 0;
      let updated = 0;
      for (const [k, v] of incoming) {
        if (!next.has(k)) added++;
        else if (next.get(k).toLowerCase() !== v.toLowerCase()) updated++;
        next.set(k, v);
      }
      setDraft(next);
      setInfo(`Imported: ${added} added, ${updated} updated, ${rejected} skipped. Click Save to keep them.`);
    } catch (e) {
      setError(e.message);
    }
  }

  async function save() {
    if (await onSave(nameMapToEntries(map), `Name map saved (${map.size} names).`)) {
      setDraft(null);
      setInfo(null);
    }
  }

  const tone = { 'NOT matched': 'danger', 'Name map - person not found': 'warning', 'Name map': 'brand', 'People sheet': 'informative', Directory: 'success' };

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <Title3>Name matching</Title3>
      <Text>The job emails a name only when it matches exactly one person. Order: shared name map, then the tracker's People sheet, then the company directory.</Text>
      {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
      {info && <MessageBar intent="info"><MessageBarBody>{info}</MessageBarBody></MessageBar>}

      <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <Dropdown placeholder="Choose a tracker" value={(trackers.find((t) => t.id === trackerId) || {}).name || ''}
          selectedOptions={[trackerId]} onOptionSelect={(_, d) => { setTrackerId(d.optionValue); setNames(null); }}>
          {trackers.map((t) => <Option key={t.id} value={t.id}>{t.name}</Option>)}
        </Dropdown>
        <Button onClick={loadNames} disabled={!trackerId || loading}>Show names in this tracker</Button>
        <span style={{ flex: 1 }} />
        <Button as="label" disabled={busy}>
          Import Excel Notifier name map...
          <input type="file" accept=".json" hidden onChange={(e) => e.target.files[0] && importFile(e.target.files[0])} />
        </Button>
        <Button appearance="primary" disabled={!draft || busy} onClick={save}>Save changes</Button>
      </div>

      {loading && <Spinner label="Reading the tracker..." />}

      {names && (
        <Table size="small">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Name in tracker</TableHeaderCell>
              <TableHeaderCell>Items</TableHeaderCell>
              <TableHeaderCell>Matched by</TableHeaderCell>
              <TableHeaderCell>Person</TableHeaderCell>
              <TableHeaderCell>Map to</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {names.map(({ name, count }) => {
              const m = matchOf(name);
              return (
                <TableRow key={name}>
                  <TableCell>{name}</TableCell>
                  <TableCell>{count}</TableCell>
                  <TableCell><Badge appearance="tint" color={tone[m.source]}>{m.source}</Badge></TableCell>
                  <TableCell>{m.user ? `${m.user.displayName} <${m.user.email}>` : (m.mapped || '-')}</TableCell>
                  <TableCell style={{ display: 'flex', gap: 6 }}>
                    <PersonPicker directory={directory} value={null} placeholder="Pick person..." onChange={(p) => setMapping(name, p)} />
                    {map.has(nameMapKey(name)) && <Button size="small" onClick={() => setMapping(name, null)}>Remove mapping</Button>}
                  </TableCell>
                </TableRow>
              );
            })}
          </TableBody>
        </Table>
      )}

      <Text weight="semibold">All saved mappings ({map.size}){draft ? ' - unsaved changes' : ''}</Text>
      {map.size > 0 && (
        <Table size="small">
          <TableBody>
            {nameMapToEntries(map).map((e) => (
              <TableRow key={e.Name}>
                <TableCell>{e.Name}</TableCell>
                <TableCell>{e.Email}</TableCell>
                <TableCell><Button size="small" onClick={() => setMapping(e.Name, null)}>Remove</Button></TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      )}
    </div>
  );
}
