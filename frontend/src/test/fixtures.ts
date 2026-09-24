import type {BundledThemeInfo, ThemeData} from '../api/types';

export const THEME: ThemeData = {
  majorFont: {latin: 'Major Sans', hangul: '메이저', symbol: 'Major Symbol'},
  minorFont: {latin: 'Minor Sans', hangul: '마이너', symbol: 'Minor Symbol'},
  monoFont: {latin: 'Mono Code', hangul: '모노'},
  options: {titleBold: true, bodyFirstLevelStyle: 'SemiBold', preserveMono: false},
};

export const PRESETS: BundledThemeInfo[] = [
  {id: 'pretendard', name: 'Pretendard', theme: THEME},
  {
    id: 'office',
    name: 'Office',
    theme: {
      majorFont: {latin: 'Aptos Display', hangul: '맑은 고딕', symbol: 'Aptos Display'},
      minorFont: {latin: 'Aptos', hangul: '맑은 고딕', symbol: 'Aptos'},
      monoFont: {latin: 'Consolas', hangul: '나눔고딕코딩'},
      options: {titleBold: true, bodyFirstLevelStyle: null, preserveMono: false},
    },
  },
];

export interface MockRoute {
  method?: string;
  path: string;
  status?: number;
  body: unknown;
  /** Return raw bytes instead of JSON. */
  isBlob?: boolean;
}

export interface RecordedRequest {
  method: string;
  path: string;
  body: BodyInit | null | undefined;
}

/** Build a fetch mock that answers the matching routes in order and records the requests. */
export function mockFetch(routes: MockRoute[]) {
  const requests: RecordedRequest[] = [];
  const remaining = [...routes];
  const fetchMock = async (input: RequestInfo | URL, init?: RequestInit) => {
    const request = input instanceof Request ? input : new Request(input, init);
    const url = new URL(request.url);
    let body: BodyInit | null | undefined;
    const contentType = request.headers.get('content-type') ?? '';
    if (contentType.startsWith('multipart/form-data')) {
      body = await request.formData();
    } else if (request.method !== 'GET') {
      body = await request.text();
    }
    requests.push({method: request.method, path: url.pathname, body});
    const index = remaining.findIndex(
      (r) => r.path === url.pathname && (r.method ?? 'GET') === request.method,
    );
    if (index < 0) {
      return new Response(JSON.stringify({detail: 'Not Found'}), {status: 404});
    }
    const [route] = remaining.splice(index, 1);
    const status = route!.status ?? 200;
    if (route!.body instanceof FormData) {
      return new Response(route!.body, {status});
    }
    if (route!.isBlob) {
      return new Response(route!.body as BodyInit, {
        status,
        headers: {'content-type': 'application/xml'},
      });
    }
    return new Response(JSON.stringify(route!.body), {
      status,
      headers: {'content-type': 'application/json'},
    });
  };
  return {fetchMock, requests};
}

export const baseRoutes = (local = false): MockRoute[] => [
  {path: '/api/config', body: {local, maxUploadSize: 1024 * 1024}},
  {path: '/api/themes', body: PRESETS},
  {path: '/api/monospace-fonts', body: ['consolas', 'fira code']},
];
