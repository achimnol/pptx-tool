import {useEffect, useState} from 'react';

import {loadStored, saveStored} from '../utils/storage';

/**
 * A state kept in localStorage under the given key. A stored value that `parse` rejects by
 * returning undefined falls back to the initial value.
 */
export function usePersistentState<T>(
  key: string,
  initial: T,
  parse: (value: unknown) => T | undefined,
): [T, (value: T) => void] {
  const [state, setState] = useState<T>(() => parse(loadStored(key)) ?? initial);
  useEffect(() => {
    saveStored(key, state);
  }, [key, state]);
  return [state, setState];
}
