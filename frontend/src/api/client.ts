import createClient from 'openapi-fetch';

import type {paths} from './schema';

export const api = createClient<paths>({
  baseUrl: window.location.origin,
  // Look up fetch on each call instead of capturing it when the module loads, so that tests can stub it.
  fetch: (request) => globalThis.fetch(request),
});
