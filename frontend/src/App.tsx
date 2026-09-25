import './App.css';

import {AppShell} from '@astryxdesign/core/AppShell';
import {Banner} from '@astryxdesign/core/Banner';
import {Heading} from '@astryxdesign/core/Heading';
import {Spinner} from '@astryxdesign/core/Spinner';
import {Tab, TabList} from '@astryxdesign/core/TabList';
import {TopNav} from '@astryxdesign/core/TopNav';
import {VStack} from '@astryxdesign/core/VStack';
import {useEffect, useMemo, useReducer, useState} from 'react';

import type {FieldError} from './api/errors';
import type {BundledThemeInfo} from './api/types';
import {FixFontsPanel} from './fix-fonts/FixFontsPanel';
import {FontThemePanel} from './font-theme/FontThemePanel';
import {useAppConfig, useBundledThemes, useMonospaceFonts} from './hooks/resources';
import {usePersistentState} from './hooks/usePersistentState';
import {ThemeEditor} from './theme-editor/ThemeEditor';
import {
  themeEditorReducer,
  type ThemeEditorAction,
  type ThemeEditorState,
} from './theme-editor/themeState';
import {validateTheme} from './theme-editor/validateTheme';
import {loadStored, saveStored} from './utils/storage';

type TaskTab = 'fix-fonts' | 'font-theme';

const DEFAULT_PRESET_ID = 'pretendard';
const EDITOR_STATE_KEY = 'themeEditor';

function parseTab(value: unknown): TaskTab | undefined {
  return value === 'fix-fonts' || value === 'font-theme' ? value : undefined;
}

type EditorAction = ThemeEditorAction | {type: 'restore'; state: ThemeEditorState};

// The editor state is empty until the bundled themes are loaded.
function editorReducer(
  state: ThemeEditorState | null,
  action: EditorAction,
): ThemeEditorState | null {
  if (action.type === 'restore') return action.state;
  if (state === null) {
    return action.type === 'selectPreset' || action.type === 'import'
      ? themeEditorReducer(
          {presetId: null, theme: action.type === 'import' ? action.theme : action.preset.theme},
          action,
        )
      : null;
  }
  return themeEditorReducer(state, action);
}

function initialPreset(presets: readonly BundledThemeInfo[]): BundledThemeInfo | undefined {
  return presets.find((p) => p.id === DEFAULT_PRESET_ID) ?? presets[0];
}

/** The editor state of the previous visit, if it holds a valid theme. */
function loadEditorState(presets: readonly BundledThemeInfo[]): ThemeEditorState | undefined {
  const stored = loadStored(EDITOR_STATE_KEY);
  if (typeof stored !== 'object' || stored === null) return undefined;
  const {presetId, theme} = stored as Record<string, unknown>;
  const validation = validateTheme(theme);
  if (!validation.ok) return undefined;
  // A preset that the server no longer bundles leaves the theme as an unnamed one.
  const isKnownPreset = typeof presetId === 'string' && presets.some((p) => p.id === presetId);
  return {presetId: isKnownPreset ? presetId : null, theme: validation.theme};
}

export function App() {
  const config = useAppConfig();
  const presets = useBundledThemes();
  const monospaceFonts = useMonospaceFonts();
  const [tab, setTab] = usePersistentState<TaskTab>('tab', 'fix-fonts', parseTab);
  const [editorState, dispatch] = useReducer(editorReducer, null);
  const [serverFieldErrors, setServerFieldErrors] = useState<FieldError[]>([]);

  useEffect(() => {
    if (editorState !== null || !presets.data) return;
    const stored = loadEditorState(presets.data);
    const preset = initialPreset(presets.data);
    if (stored) {
      dispatch({type: 'restore', state: stored});
    } else if (preset) {
      dispatch({type: 'selectPreset', preset});
    }
  }, [editorState, presets.data]);

  useEffect(() => {
    if (editorState !== null) saveStored(EDITOR_STATE_KEY, editorState);
  }, [editorState]);

  const validation = useMemo(
    () => (editorState ? validateTheme(editorState.theme) : undefined),
    [editorState],
  );
  const fieldErrors = validation && !validation.ok ? validation.errors : serverFieldErrors;
  const loadError = config.error ?? presets.error ?? monospaceFonts.error;

  let content;
  if (loadError) {
    content = (
      <Banner
        status="error"
        title="Failed to connect to the server"
        description={loadError.message}
      />
    );
  } else if (editorState === null) {
    content = <Spinner label="Loading" />;
  } else {
    content = (
      <div className="app-columns">
        <ThemeEditor
          state={editorState}
          dispatch={(action) => {
            setServerFieldErrors([]);
            dispatch(action);
          }}
          presets={presets.data ?? []}
          monospaceFonts={monospaceFonts.data ?? []}
          fieldErrors={fieldErrors}
        />
        <VStack gap={6}>
          <Heading level={2}>Operations</Heading>
          <TabList
            value={tab}
            onChange={(value) => setTab(value as TaskTab)}
            role="tablist"
            hasDivider
          >
            <Tab value="fix-fonts" label="Fix fonts" />
            <Tab value="font-theme" label="Export office font theme" />
          </TabList>
          {/* Keep both task panels mounted so that switching tabs keeps their inputs. */}
          <div hidden={tab !== 'fix-fonts'}>
            <FixFontsPanel
              theme={editorState.theme}
              isThemeValid={validation?.ok ?? false}
              maxUploadSize={config.data?.maxUploadSize}
              onFieldErrors={setServerFieldErrors}
            />
          </div>
          <div hidden={tab !== 'font-theme'}>
            <FontThemePanel
              theme={editorState.theme}
              isThemeValid={validation?.ok ?? false}
              isLocal={config.data?.local ?? false}
              onFieldErrors={setServerFieldErrors}
            />
          </div>
        </VStack>
      </div>
    );
  }

  return (
    <AppShell
      topNav={<TopNav className="app-topnav" heading={<Heading level={1}>pptx-tool</Heading>} />}
      contentPadding={6}
    >
      <div className="app-content">{content}</div>
    </AppShell>
  );
}
