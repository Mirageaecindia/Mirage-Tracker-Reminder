// ================================================================
// web/src/components/PersonPicker.jsx
// ----------------------------------------------------------------
// PURPOSE : Type-to-search box over the company directory; returns ONE
//           person { name, email }. Used for lead, on-behalf person, admins
//           and name matching. Only UI (SRP).
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Searches display name AND email, case-insensitive; shows at most 30
//       matches (the list has ~110 people; 30 keeps the dropdown short).
//     * Only people FROM the directory can be picked (no free typing):
//       the directory is already limited to @mirageaec.com addresses, so a
//       pick can never be an outside address.
// ================================================================

import { useMemo, useState } from 'react';
import { Combobox, Option } from '@fluentui/react-components';

const MAX_SHOWN = 30;

/**
 * @param {{ directory: {displayName: string, email: string}[],
 *           value: {name: string, email: string} | null,
 *           onChange: (person: {name: string, email: string} | null) => void,
 *           placeholder?: string, disabled?: boolean }} props
 */
export default function PersonPicker({ directory, value, onChange, placeholder, disabled }) {
  const [query, setQuery] = useState('');

  const matches = useMemo(() => {
    const q = query.trim().toLowerCase();
    const list = q
      ? directory.filter((u) => u.displayName.toLowerCase().includes(q) || u.email.toLowerCase().includes(q))
      : directory;
    return list.slice(0, MAX_SHOWN);
  }, [directory, query]);

  return (
    <Combobox
      disabled={disabled}
      placeholder={placeholder || 'Type a name...'}
      value={query || (value ? `${value.name} <${value.email}>` : '')}
      selectedOptions={value ? [value.email] : []}
      onChange={(e) => setQuery(e.target.value)}
      onOptionSelect={(_, data) => {
        const user = directory.find((u) => u.email === data.optionValue);
        onChange(user ? { name: user.displayName, email: user.email } : null);
        setQuery('');
      }}
      style={{ minWidth: 320 }}
    >
      {matches.map((u) => (
        <Option key={u.email} value={u.email} text={`${u.displayName} <${u.email}>`}>
          {u.displayName} &lt;{u.email}&gt;
        </Option>
      ))}
      {matches.length === 0 && <Option disabled value="" text="No match">No match</Option>}
    </Combobox>
  );
}
