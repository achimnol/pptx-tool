import {useEffect, useState} from 'react';

export interface ApiResource<T> {
  data: T | undefined;
  error: Error | undefined;
  isLoading: boolean;
}

/** Load a resource once when the component mounts. */
export function useApiResource<T>(load: () => Promise<T>): ApiResource<T> {
  const [state, setState] = useState<ApiResource<T>>({
    data: undefined,
    error: undefined,
    isLoading: true,
  });
  useEffect(() => {
    let isCancelled = false;
    load().then(
      (data) => !isCancelled && setState({data, error: undefined, isLoading: false}),
      (error: unknown) =>
        !isCancelled &&
        setState({
          data: undefined,
          error: error instanceof Error ? error : new Error(String(error)),
          isLoading: false,
        }),
    );
    return () => {
      isCancelled = true;
    };
    // The loader is expected to be a module-level function.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  return state;
}
