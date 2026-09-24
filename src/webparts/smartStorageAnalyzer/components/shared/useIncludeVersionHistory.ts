import * as React from 'react';
import { safeGet, safeSet } from '../../utils/safeStorage';

// Backs the Storage Report's "Include Version History Size" checkbox,
// remembered both for the rest of the session and across future ones. Tree
// View/List View always shows version history (it's fetched at no extra
// cost there regardless), so this only matters for the Storage Report,
// where it's real extra measurement work on some libraries. Read as "on
// unless explicitly turned off", so a first-ever run (nothing stored yet)
// defaults on — Version History Size is what makes Total Storage Size the
// true storage number.
const LS_KEY = 'sp-smart-storage-analyzer-includeVersionHistory';

export function useIncludeVersionHistory(): [boolean, (value: boolean) => void] {
  const [value, setValue] = React.useState(() => safeGet(LS_KEY) !== 'false');
  React.useEffect(() => {
    safeSet(LS_KEY, String(value));
  }, [value]);
  return [value, setValue];
}
