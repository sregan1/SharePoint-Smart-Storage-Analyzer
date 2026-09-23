// localStorage throws (SecurityError) when storage is blocked — Teams and
// partitioned iframes, strict privacy settings — and setItem throws when it's
// full. A preference that can't be read or saved must never take the web part
// down with it, so every access goes through these and degrades to "no saved
// value" instead.

export function safeGet(key: string): string | null {
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

export function safeSet(key: string, value: string): void {
  try {
    localStorage.setItem(key, value);
  } catch {
    // Not persisted this time; the in-memory value still applies.
  }
}
