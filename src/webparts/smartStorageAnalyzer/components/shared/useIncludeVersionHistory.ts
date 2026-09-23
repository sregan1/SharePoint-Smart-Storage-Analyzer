import * as React from 'react';

// Shared across every screen (Tree View, List View, Storage Report) so
// toggling "Include Version History Size" anywhere is remembered everywhere,
// both for the rest of the session and across future ones. Stored as "was it
// explicitly turned off", not "was it turned on" — so a first-ever run
// (nothing in localStorage yet) still gets the true default, and only an
// explicit uncheck is remembered as false.
const LS_KEY = 'sp-smart-storage-analyzer-includeVersionHistory';

export function useIncludeVersionHistory(): [boolean, (value: boolean) => void] {
  const [value, setValue] = React.useState(() => localStorage.getItem(LS_KEY) !== 'false');
  React.useEffect(() => {
    localStorage.setItem(LS_KEY, String(value));
  }, [value]);
  return [value, setValue];
}
