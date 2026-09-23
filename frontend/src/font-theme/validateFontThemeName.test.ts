import {describe, expect, it} from 'vitest';

import {validateFontThemeName} from './validateFontThemeName';

describe('validateFontThemeName', () => {
  it.each(['My Theme', '나의 테마', '  padded  ', 'a.b', 'x'.repeat(100)])('accepts %j', (name) => {
    expect(validateFontThemeName(name)).toBeNull();
  });

  it.each([
    '',
    '   ',
    '.',
    '..',
    '../evil',
    'ends.',
    'a/b',
    'a\\b',
    'C:',
    'a*b',
    'a?b',
    'a"b',
    'a<b',
    'a|b',
    'a\nb',
    'x'.repeat(101),
  ])('rejects %j', (name) => {
    expect(validateFontThemeName(name)).not.toBeNull();
  });
});
