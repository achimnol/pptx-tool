import {describe, expect, it} from 'vitest';

import {THEME} from '../test/fixtures';
import {validateTheme} from './validateTheme';

function withValue(path: string, value: unknown): unknown {
  const data = structuredClone(THEME) as unknown as Record<string, Record<string, unknown>>;
  const [parent, key] = path.split('.') as [string, string];
  if (value === undefined) {
    delete data[parent]![key];
  } else {
    data[parent]![key] = value;
  }
  return data;
}

describe('validateTheme', () => {
  it('accepts a valid theme', () => {
    expect(validateTheme(THEME)).toEqual({ok: true, theme: THEME});
  });

  it('defaults preserveMono to false', () => {
    const result = validateTheme(withValue('options.preserveMono', undefined));
    expect(result.ok && result.theme.options.preserveMono).toBe(false);
  });

  it.each([
    ['majorFont.latin', ''],
    ['majorFont.latin', '  '],
    ['minorFont.symbol', 3],
    ['monoFont.hangul', undefined],
    ['options.titleBold', 'true'],
    ['options.bodyFirstLevelStyle', undefined],
    ['options.bodyFirstLevelStyle', ''],
    ['options.preserveMono', 1],
  ])('reports %s = %j', (path, value) => {
    const result = validateTheme(withValue(path, value));
    expect(result.ok).toBe(false);
    expect(!result.ok && result.errors.map((e) => e.key)).toEqual([path]);
  });

  it('accepts a null bodyFirstLevelStyle', () => {
    expect(validateTheme(withValue('options.bodyFirstLevelStyle', null)).ok).toBe(true);
  });

  it('rejects non-objects', () => {
    expect(validateTheme([])).toEqual({
      ok: false,
      errors: [{key: '', message: 'the theme must be a JSON object'}],
    });
  });
});
