// The user choices kept in the browser's localStorage across visits.
// Bump the version when a stored shape changes incompatibly, so that old values are ignored.
const PREFIX = 'pptx-tool:v1:';

/** Read a stored value, or undefined when it is missing, malformed or the storage is unavailable. */
export function loadStored(key: string): unknown {
  try {
    const raw = window.localStorage.getItem(PREFIX + key);
    return raw === null ? undefined : (JSON.parse(raw) as unknown);
  } catch {
    return undefined;
  }
}

/** Store a value, ignoring failures such as a disabled storage or an exceeded quota. */
export function saveStored(key: string, value: unknown): void {
  try {
    window.localStorage.setItem(PREFIX + key, JSON.stringify(value));
  } catch {
    // The choices are a convenience; the app works the same without them.
  }
}
