// ================================================================
// web/src/App.jsx
// ----------------------------------------------------------------
// PURPOSE : Page frame + the ONE owner of the shared settings files:
//             - header (Mirage navy / gold, logo, signed-in person, sign out)
//             - sign-in screen, admin check (admins.json)
//             - loads trackers.json, name_map.json, admins.json (with their
//               eTags), runstate.json, sentlog.json and the company directory
//             - tabs: Trackers / Name matching / Preview / Admins
//             - every SAVE goes through here, so each file has exactly one
//               current eTag (no two screens overwriting each other)
//           Screens themselves are in pages/ (SRP).
//
// CHANGE LOG
// - 2026-10-14 : Initial version (user decisions 2026-10-14: four screens,
//       Fluent look with Mirage colours, only listed admins).
//     * Admin check = signed-in email in admins.json. NOTE (told to the user):
//       this only gates the PAGE; the real protection is SharePoint edit
//       rights on the Automation site - keep them to the same admins.
//     * A missing admins.json shows how to create it (upload
//       config-samples/admins.json) instead of letting everyone in.
//     * Save conflicts (someone else saved first) -> message + reload, never
//       a silent overwrite.
// ================================================================

import { useCallback, useEffect, useMemo, useState } from 'react';
import { useIsAuthenticated, useMsal } from '@azure/msal-react';
import {
  Button, MessageBar, MessageBarBody, Spinner, Tab, TabList, Text, Title2,
} from '@fluentui/react-components';

import { loginRequest } from './authConfig.js';
import { NAVY, GOLD } from './theme.js';
import { Graph, ConflictError } from './services/graph.js';
import { ConfigStore, FILES, adminEmails } from './services/configStore.js';
import { directoryFromGraphUsers } from '../../core/recipientResolver.js';
import logoUrl from '../../assets/mirage_logo.jpg';

import TrackersPage from './pages/TrackersPage.jsx';
import TrackerEditor from './pages/TrackerEditor.jsx';
import NamesPage from './pages/NamesPage.jsx';
import PreviewPage from './pages/PreviewPage.jsx';
import AdminsPage from './pages/AdminsPage.jsx';

export default function App() {
  const { instance } = useMsal();
  const signedIn = useIsAuthenticated();

  if (!signedIn) {
    return (
      <Shell>
        <div style={{ maxWidth: 520, margin: '64px auto', textAlign: 'center' }}>
          <Title2>Tracker Reminder - admin</Title2>
          <p><Text>Sign in with your Mirage Microsoft account to manage tracker reminders.</Text></p>
          <Button appearance="primary" size="large" onClick={() => instance.loginRedirect(loginRequest)}>
            Sign in with Microsoft
          </Button>
        </div>
      </Shell>
    );
  }
  return <AdminApp msal={instance} />;
}

function AdminApp({ msal }) {
  const graph = useMemo(() => new Graph(msal), [msal]);
  const store = useMemo(() => new ConfigStore(graph), [graph]);

  const [me, setMe] = useState(null);
  const [access, setAccess] = useState('loading'); // loading | ok | denied | missing
  const [error, setError] = useState(null);
  const [notice, setNotice] = useState(null);
  const [busy, setBusy] = useState(false);

  const [directory, setDirectory] = useState([]);
  const [trackers, setTrackers] = useState({ data: null, eTag: null });
  const [admins, setAdmins] = useState({ data: null, eTag: null });
  const [nameMap, setNameMap] = useState({ data: null, eTag: null });
  const [runState, setRunState] = useState(null);
  const [sentLog, setSentLog] = useState(null);

  const [tab, setTab] = useState('trackers');
  const [editingId, setEditingId] = useState(undefined); // undefined = list, null = new tracker, id = edit

  const loadAll = useCallback(async () => {
    setError(null);
    try {
      const [who, adminsDoc] = await Promise.all([graph.me(), store.load(FILES.admins)]);
      const email = String(who.mail || who.userPrincipalName || '').toLowerCase();
      setMe({ name: who.displayName, email });
      setAdmins(adminsDoc);
      if (!adminsDoc.data) return setAccess('missing');
      if (!adminEmails(adminsDoc.data).has(email)) return setAccess('denied');

      const [users, trackersDoc, nameMapDoc, runDoc, sentDoc] = await Promise.all([
        graph.listUsers(), store.load(FILES.trackers), store.load(FILES.nameMap),
        store.load(FILES.runState), store.load(FILES.sentLog),
      ]);
      setDirectory(directoryFromGraphUsers(users));
      setTrackers(trackersDoc);
      setNameMap(nameMapDoc);
      setRunState(runDoc.data);
      setSentLog(sentDoc.data);
      setAccess('ok');
    } catch (e) {
      setError(e.message);
    }
  }, [graph, store]);

  useEffect(() => { loadAll(); }, [loadAll]);

  /** One save path for every file (see change log). */
  const save = useCallback(async (name, data, current, setDoc, message) => {
    setBusy(true);
    setError(null);
    try {
      const eTag = await store.save(name, data, current.eTag);
      setDoc({ data, eTag });
      setNotice(message);
      return true;
    } catch (e) {
      setError(e instanceof ConflictError ? `${e.message} (${name})` : e.message);
      if (e instanceof ConflictError) await loadAll();
      return false;
    } finally {
      setBusy(false);
    }
  }, [store, loadAll]);

  const saveTrackers = (data, message) => save(FILES.trackers, data, trackers, setTrackers, message);
  const saveNameMap = (data, message) => save(FILES.nameMap, data, nameMap, setNameMap, message);
  const saveAdmins = (data, message) => save(FILES.admins, data, admins, setAdmins, message);

  const trackerList = trackers.data && Array.isArray(trackers.data.trackers) ? trackers.data.trackers : [];

  return (
    <Shell me={me} onSignOut={() => msal.logoutRedirect()}>
      {error && <MessageBar intent="error" style={{ marginBottom: 12 }}><MessageBarBody>{error}</MessageBarBody></MessageBar>}
      {notice && <MessageBar intent="success" style={{ marginBottom: 12 }}><MessageBarBody>{notice}</MessageBarBody></MessageBar>}

      {access === 'loading' && !error && <Spinner label="Loading settings from SharePoint..." />}
      {access === 'missing' && (
        <MessageBar intent="warning"><MessageBarBody>
          admins.json was not found in Automation / Documents / Application_Files / TrackerReminder.
          Upload config-samples\admins.json there, then reload this page.
        </MessageBarBody></MessageBar>
      )}
      {access === 'denied' && (
        <MessageBar intent="error"><MessageBarBody>
          {me ? me.email : 'This account'} is not on the admin list. Ask an admin to add you on the Admins tab.
        </MessageBarBody></MessageBar>
      )}

      {access === 'ok' && (
        <>
          <TabList selectedValue={tab} onTabSelect={(_, d) => { setTab(d.value); setEditingId(undefined); setNotice(null); }} style={{ marginBottom: 16 }}>
            <Tab value="trackers">Trackers</Tab>
            <Tab value="names">Name matching</Tab>
            <Tab value="preview">Preview</Tab>
            <Tab value="admins">Admins</Tab>
          </TabList>

          {tab === 'trackers' && editingId === undefined && (
            <TrackersPage
              trackersData={trackers.data}
              runState={runState}
              busy={busy}
              onAdd={() => { setNotice(null); setEditingId(null); }}
              onEdit={(id) => { setNotice(null); setEditingId(id); }}
              onToggleEnabled={(id, enabled) => saveTrackers(
                { version: 1, ...(trackers.data || {}), trackers: trackerList.map((t) => (t.id === id ? { ...t, enabled } : t)) },
                `${enabled ? 'Enabled' : 'Disabled'} - saved.`,
              )}
            />
          )}
          {tab === 'trackers' && editingId !== undefined && (
            <TrackerEditor
              graph={graph}
              directory={directory}
              me={me}
              trackers={trackerList}
              editing={editingId === null ? null : trackerList.find((t) => t.id === editingId)}
              busy={busy}
              onCancel={() => setEditingId(undefined)}
              onSave={async (tracker) => {
                const others = trackerList.filter((t) => t.id !== tracker.id);
                const ok = await saveTrackers({ version: 1, ...(trackers.data || {}), trackers: [...others, tracker] }, `"${tracker.name}" saved.`);
                if (ok) setEditingId(undefined);
              }}
              onDelete={async (id) => {
                const ok = await saveTrackers({ version: 1, ...(trackers.data || {}), trackers: trackerList.filter((t) => t.id !== id) }, 'Tracker removed.');
                if (ok) setEditingId(undefined);
              }}
            />
          )}
          {tab === 'names' && (
            <NamesPage
              graph={graph}
              directory={directory}
              trackers={trackerList}
              nameMapData={nameMap.data}
              busy={busy}
              onSave={(entries, message) => saveNameMap(entries, message)}
            />
          )}
          {tab === 'preview' && (
            <PreviewPage
              graph={graph}
              directory={directory}
              trackers={trackerList}
              nameMapData={nameMap.data}
              sentLogData={sentLog}
            />
          )}
          {tab === 'admins' && (
            <AdminsPage
              directory={directory}
              adminsData={admins.data}
              me={me}
              busy={busy}
              onSave={(data, message) => saveAdmins(data, message)}
            />
          )}
        </>
      )}
    </Shell>
  );
}

/** Header bar + page body. */
function Shell({ me, onSignOut, children }) {
  return (
    <div>
      <header style={{
        background: NAVY, borderBottom: `3px solid ${GOLD}`, color: '#fff',
        display: 'flex', alignItems: 'center', gap: 16, padding: '10px 24px',
      }}>
        <img src={logoUrl} alt="Mirage AEC" style={{ height: 36 }} />
        <span style={{ fontSize: 18, fontWeight: 600, letterSpacing: 0.3 }}>Tracker Reminder</span>
        <span style={{ flex: 1 }} />
        {me && <span style={{ fontSize: 13, opacity: 0.85 }}>{me.name} &lt;{me.email}&gt;</span>}
        {onSignOut && <Button size="small" onClick={onSignOut}>Sign out</Button>}
      </header>
      <main style={{ padding: '20px 24px', maxWidth: 1280, margin: '0 auto' }}>{children}</main>
    </div>
  );
}
