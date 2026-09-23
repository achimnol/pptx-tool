import type {FieldError} from '../api/errors';
import type {ThemeData} from '../api/types';
import {FONT_SETS} from './fields';

export type ThemeValidation = {ok: true; theme: ThemeData} | {ok: false; errors: FieldError[]};

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.trim() !== '';
}

/**
 * Validate a theme definition with the same rules and field paths as the server
 * (`pptx_tool.theme.load_theme()`), returning a normalized copy when valid.
 */
export function validateTheme(data: unknown): ThemeValidation {
  if (!isObject(data)) {
    return {ok: false, errors: [{key: '', message: 'the theme must be a JSON object'}]};
  }
  const errors: FieldError[] = [];
  const getObject = (parent: Record<string, unknown>, key: string, path: string) => {
    const value = parent[key];
    if (!isObject(value)) {
      errors.push({key: path, message: 'must be an object'});
      return {};
    }
    return value;
  };
  const fonts: Record<string, Record<string, string>> = {};
  for (const {key: setKey, scripts} of FONT_SETS) {
    const fontSet = getObject(data, setKey, setKey);
    fonts[setKey] = {};
    for (const script of scripts) {
      const value = fontSet[script];
      if (isNonEmptyString(value)) {
        fonts[setKey][script] = value;
      } else {
        errors.push({key: `${setKey}.${script}`, message: 'must be a non-empty string'});
      }
    }
  }
  const options = getObject(data, 'options', 'options');
  const {titleBold, bodyFirstLevelStyle, preserveMono = false} = options;
  if (typeof titleBold !== 'boolean') {
    errors.push({key: 'options.titleBold', message: 'must be a boolean'});
  }
  if (!('bodyFirstLevelStyle' in options)) {
    errors.push({
      key: 'options.bodyFirstLevelStyle',
      message: 'must be present (use null to disable it)',
    });
  } else if (bodyFirstLevelStyle !== null && !isNonEmptyString(bodyFirstLevelStyle)) {
    errors.push({
      key: 'options.bodyFirstLevelStyle',
      message: 'must be a non-empty string or null',
    });
  }
  if (typeof preserveMono !== 'boolean') {
    errors.push({key: 'options.preserveMono', message: 'must be a boolean'});
  }
  if (errors.length > 0) {
    return {ok: false, errors};
  }
  return {
    ok: true,
    theme: {
      majorFont: fonts.majorFont as ThemeData['majorFont'],
      minorFont: fonts.minorFont as ThemeData['minorFont'],
      monoFont: fonts.monoFont as ThemeData['monoFont'],
      options: {
        titleBold: titleBold as boolean,
        bodyFirstLevelStyle: bodyFirstLevelStyle as string | null,
        preserveMono: preserveMono as boolean,
      },
    },
  };
}
