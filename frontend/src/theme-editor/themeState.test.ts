import {describe, expect, it} from 'vitest';

import {PRESETS, THEME} from '../test/fixtures';
import {isModified, themeEditorReducer, type ThemeEditorState} from './themeState';

const initial: ThemeEditorState = {presetId: 'pretendard', theme: THEME};

describe('themeEditorReducer', () => {
  it('updates fonts and options immutably', () => {
    let state = themeEditorReducer(initial, {
      type: 'setFont',
      set: 'minorFont',
      script: 'hangul',
      value: '새 글꼴',
    });
    state = themeEditorReducer(state, {type: 'setPreserveMono', value: true});
    state = themeEditorReducer(state, {type: 'setBodyFirstLevelStyle', value: null});
    state = themeEditorReducer(state, {type: 'setTitleBold', value: false});
    expect(state.theme.minorFont.hangul).toBe('새 글꼴');
    expect(state.theme.options).toEqual({
      titleBold: false,
      bodyFirstLevelStyle: null,
      preserveMono: true,
    });
    expect(THEME.minorFont.hangul).toBe('마이너');
  });

  it('selects presets and imports themes', () => {
    const office = PRESETS[1]!;
    expect(themeEditorReducer(initial, {type: 'selectPreset', preset: office})).toEqual({
      presetId: 'office',
      theme: office.theme,
    });
    expect(themeEditorReducer(initial, {type: 'import', theme: office.theme}).presetId).toBeNull();
  });
});

describe('isModified', () => {
  it('compares the theme with its preset', () => {
    expect(isModified(initial, PRESETS)).toBe(false);
    const edited = themeEditorReducer(initial, {type: 'setTitleBold', value: false});
    expect(isModified(edited, PRESETS)).toBe(true);
    const reverted = themeEditorReducer(edited, {type: 'setTitleBold', value: true});
    expect(isModified(reverted, PRESETS)).toBe(false);
    expect(isModified({presetId: null, theme: THEME}, PRESETS)).toBe(false);
  });
});
