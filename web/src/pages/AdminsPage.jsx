// ================================================================
// web/src/pages/AdminsPage.jsx
// ----------------------------------------------------------------
// PURPOSE : Who may use this page (admins.json). Add from the company
//           directory, remove. Only UI (SRP); App.jsx saves.
//
// CHANGE LOG
// - 2026-10-14 : Initial version (user decision: only listed admins; first
//       admins = Vasanthan + mirage.automation).
//     * You cannot remove YOURSELF (you would lock yourself out mid-session)
//       nor the LAST admin (nobody could get back in without editing the
//       JSON file by hand).
//     * Reminder shown on screen: this list only gates the PAGE - SharePoint
//       edit rights on the Automation site are the real protection.
// ================================================================

import { useState } from 'react';
import {
  Button, MessageBar, MessageBarBody, Table, TableBody, TableCell, TableHeader,
  TableHeaderCell, TableRow, Text, Title3,
} from '@fluentui/react-components';

import PersonPicker from '../components/PersonPicker.jsx';

/**
 * @param {{ directory: object[], adminsData: {admins: {name: string, email: string}[]},
 *           me: {email: string}, busy: boolean, onSave: (data: object, message: string) => Promise<boolean> }} props
 */
export default function AdminsPage({ directory, adminsData, me, busy, onSave }) {
  const admins = (adminsData && Array.isArray(adminsData.admins) ? adminsData.admins : []);
  const [toAdd, setToAdd] = useState(null);

  const isMe = (a) => a.email.toLowerCase() === me.email.toLowerCase();

  function add() {
    if (!toAdd || admins.some((a) => a.email.toLowerCase() === toAdd.email.toLowerCase())) return;
    onSave({ version: 1, admins: [...admins, toAdd] }, `${toAdd.name} is now an admin.`).then((ok) => ok && setToAdd(null));
  }

  function remove(a) {
    onSave({ version: 1, admins: admins.filter((x) => x.email.toLowerCase() !== a.email.toLowerCase()) }, `${a.name} removed.`);
  }

  return (
    <div style={{ display: 'grid', gap: 16, maxWidth: 800 }}>
      <Title3>Admins</Title3>
      <MessageBar intent="info"><MessageBarBody>
        This list decides who can open this page. The settings files themselves are protected by SharePoint:
        keep edit access to the Automation site limited to these same people.
      </MessageBarBody></MessageBar>

      <Table size="small">
        <TableHeader>
          <TableRow><TableHeaderCell>Name</TableHeaderCell><TableHeaderCell>Email</TableHeaderCell><TableHeaderCell /></TableRow>
        </TableHeader>
        <TableBody>
          {admins.map((a) => (
            <TableRow key={a.email}>
              <TableCell>{a.name}{isMe(a) ? ' (you)' : ''}</TableCell>
              <TableCell>{a.email}</TableCell>
              <TableCell>
                <Button size="small" disabled={busy || isMe(a) || admins.length <= 1} onClick={() => remove(a)}>Remove</Button>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>

      <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <PersonPicker directory={directory} value={toAdd} onChange={setToAdd} placeholder="Add an admin..." />
        <Button appearance="primary" onClick={add} disabled={!toAdd || busy}>Add</Button>
      </div>
      <Text size={200}>You cannot remove yourself or the last admin.</Text>
    </div>
  );
}
