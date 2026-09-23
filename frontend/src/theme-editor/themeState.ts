import type {BundledThemeInfo, ThemeData} from '../api/types';
import type {FontSetKey, ScriptKey} from './fields';

export interface ThemeEditorState {
  /** The bundled theme the editor started from, or null for an imported theme. */
  presetId: string | null;
  theme: ThemeData;
}

export type ThemeEditorAction =
  | {type: 'selectPreset'; preset: BundledThemeInfo}
  | {type: 'import'; theme: ThemeData}
  | {type: 'setFont'; set: FontSetKey; script: ScriptKey; value: string}
  | {type: 'setTitleBold'; value: boolean}
  | {type: 'setBodyFirstLevelStyle'; value: string | null}
  | {type: 'setPreserveMono'; value: boolean};

export function themeEditorReducer(
  state: ThemeEditorState,
  action: ThemeEditorAction,
): ThemeEditorState {
  const {theme} = state;
  switch (action.type) {
    case 'selectPreset':
      return {presetId: action.preset.id, theme: action.preset.theme};
    case 'import':
      return {presetId: null, theme: action.theme};
    case 'setFont':
      return {
        ...state,
        theme: {...theme, [action.set]: {...theme[action.set], [action.script]: action.value}},
      };
    case 'setTitleBold':
      return {...state, theme: {...theme, options: {...theme.options, titleBold: action.value}}};
    case 'setBodyFirstLevelStyle':
      return {
        ...state,
        theme: {...theme, options: {...theme.options, bodyFirstLevelStyle: action.value}},
      };
    case 'setPreserveMono':
      return {...state, theme: {...theme, options: {...theme.options, preserveMono: action.value}}};
  }
}

/** Whether the theme differs from the preset it started from. */
export function isModified(state: ThemeEditorState, presets: readonly BundledThemeInfo[]): boolean {
  const preset = presets.find((p) => p.id === state.presetId);
  if (preset === undefined) return false;
  return JSON.stringify(preset.theme) !== JSON.stringify(state.theme);
}
