import {api} from '../api/client';
import {toApiError} from '../api/errors';
import type {AppConfig, BundledThemeInfo} from '../api/types';
import {useApiResource} from './useApiResource';

async function loadAppConfig(): Promise<AppConfig> {
  const {data, error, response} = await api.GET('/api/config');
  if (data === undefined) throw toApiError(response, error);
  return data;
}

async function loadBundledThemes(): Promise<BundledThemeInfo[]> {
  const {data, error, response} = await api.GET('/api/themes');
  if (data === undefined) throw toApiError(response, error);
  return data;
}

async function loadMonospaceFonts(): Promise<string[]> {
  const {data, error, response} = await api.GET('/api/monospace-fonts');
  if (data === undefined) throw toApiError(response, error);
  return data;
}

export const useAppConfig = () => useApiResource(loadAppConfig);
export const useBundledThemes = () => useApiResource(loadBundledThemes);
export const useMonospaceFonts = () => useApiResource(loadMonospaceFonts);
