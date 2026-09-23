import {api} from './client';
import {toApiError} from './errors';
import type {FixFontResult, ThemeData} from './types';

export async function fixFont(file: File, theme: ThemeData): Promise<FixFontResult> {
  const {data, error, response} = await api.POST('/api/fix-font', {
    // The theme travels as a JSON-encoded form field next to the file.
    body: {file: file as unknown as string, theme: JSON.stringify(theme)},
    bodySerializer: (body) => {
      const form = new FormData();
      form.append('file', body.file as unknown as File);
      form.append('theme', body.theme);
      return form;
    },
  });
  if (data === undefined) throw toApiError(response, error);
  return data;
}

export async function downloadFontTheme(name: string, theme: ThemeData): Promise<Blob> {
  const {data, error, response} = await api.POST('/api/font-theme', {
    body: {name, theme},
    parseAs: 'blob',
  });
  if (data === undefined) throw toApiError(response, error);
  return data as Blob;
}

export async function installFontTheme(
  name: string,
  theme: ThemeData,
  overwrite: boolean,
): Promise<string> {
  const {data, error, response} = await api.POST('/api/font-theme/install', {
    body: {name, theme, overwrite},
  });
  if (data === undefined) throw toApiError(response, error);
  return data.path;
}
