import * as React from 'react';
import { safeGet, safeSet } from '../../utils/safeStorage';

// Shared across every screen (Tree View, List View, Storage Report) so
// toggling "Include Version History Size" anywhere is remembered everywhere,
// both for the rest of the session and across future ones. Read as "off
// unless explicitly turned on", so a first-ever run (nothing stored yet)
// defaults off — measuring version history is real extra work (a bulk sweep
// at minimum, a per-file pass on some libraries), and defaulting it on made
// every first scan/folder load slower than necessary.
const LS_KEY = 'sp-smart-storage-analyzer-includeVersionHistory';

export function useIncludeVersionHistory(): [boolean, (value: boolean) => void] {
  const [value, setValue] = React.useState(() => safeGet(LS_KEY) === 'true');
  React.useEffect(() => {
    safeSet(LS_KEY, String(value));
  }, [value]);
  return [value, setValue];
}
