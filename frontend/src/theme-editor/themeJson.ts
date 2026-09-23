import type {ThemeData} from '../api/types';
import {validateTheme, type ThemeValidation} from './validateTheme';

/** Serialize the theme as a JSON file usable with the CLI's --theme option. */
export function exportTheme(theme: ThemeData): Blob {
  return new Blob([JSON.stringify(theme, null, 2) + '\n'], {type: 'application/json'});
}

export async function parseThemeFile(file: File): Promise<ThemeValidation> {
  let data: unknown;
  try {
    data = JSON.parse(await file.text());
  } catch (e) {
    return {ok: false, errors: [{key: '', message: `Not a valid JSON file: ${String(e)}`}]};
  }
  return validateTheme(data);
}
