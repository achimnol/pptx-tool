import {Button} from '@astryxdesign/core/Button';
import {FileInput} from '@astryxdesign/core/FileInput';
import {FormLayout} from '@astryxdesign/core/FormLayout';
import {Grid} from '@astryxdesign/core/Grid';
import {Heading} from '@astryxdesign/core/Heading';
import {HStack} from '@astryxdesign/core/HStack';
import {Selector} from '@astryxdesign/core/Selector';
import {Switch} from '@astryxdesign/core/Switch';
import {Text} from '@astryxdesign/core/Text';
import {TextInput} from '@astryxdesign/core/TextInput';
import {VStack} from '@astryxdesign/core/VStack';
import {useState, type Dispatch} from 'react';

import type {FieldError} from '../api/errors';
import type {BundledThemeInfo} from '../api/types';
import {MonospaceFontsPopover} from '../components/MonospaceFontsPopover';
import {downloadBlob} from '../utils/download';
import {
  BODY_FIRST_LEVEL_STYLES,
  FONT_SETS,
  SCRIPT_LABELS,
  fontPath,
  getFont,
  type ScriptKey,
} from './fields';
import {exportTheme, parseThemeFile} from './themeJson';
import {isModified, type ThemeEditorAction, type ThemeEditorState} from './themeState';

const IMPORTED_PRESET = '__imported__';
const NO_STYLE = '__none__';
const GRID_SCRIPTS: readonly ScriptKey[] = ['latin', 'hangul', 'symbol'];

export interface ThemeEditorProps {
  state: ThemeEditorState;
  dispatch: Dispatch<ThemeEditorAction>;
  presets: readonly BundledThemeInfo[];
  monospaceFonts: readonly string[];
  /** Field errors to show, keyed by the theme JSON path such as "majorFont.latin". */
  fieldErrors: readonly FieldError[];
}

function errorStatus(errors: readonly FieldError[], key: string) {
  const error = errors.find((e) => e.key === key);
  return error ? {type: 'error' as const, message: error.message} : undefined;
}

export function ThemeEditor({
  state,
  dispatch,
  presets,
  monospaceFonts,
  fieldErrors,
}: ThemeEditorProps) {
  const {theme} = state;
  const [importFile, setImportFile] = useState<File | null>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const preserveMono = theme.options.preserveMono;
  const modified = isModified(state, presets);

  const presetOptions = presets.map((p) => ({
    value: p.id,
    label: p.id === state.presetId && modified ? `${p.name} (modified)` : p.name,
  }));
  if (state.presetId === null) {
    presetOptions.unshift({value: IMPORTED_PRESET, label: 'Imported theme'});
  }

  const bodyStyle = theme.options.bodyFirstLevelStyle;
  const styleOptions = [
    {value: NO_STYLE, label: 'None'},
    ...BODY_FIRST_LEVEL_STYLES.map((s) => ({value: s, label: s})),
  ];
  if (bodyStyle !== null && !styleOptions.some((o) => o.value === bodyStyle)) {
    styleOptions.push({value: bodyStyle, label: bodyStyle});
  }

  const handleImport = async (file: File | File[] | null) => {
    const selected = Array.isArray(file) ? (file[0] ?? null) : file;
    setImportFile(selected);
    setImportError(null);
    if (selected === null) return;
    const result = await parseThemeFile(selected);
    if (result.ok) {
      dispatch({type: 'import', theme: result.theme});
    } else {
      setImportError(
        result.errors.map((e) => (e.key ? `${e.key}: ${e.message}` : e.message)).join('; '),
      );
    }
  };

  const handleExport = () => {
    const name = presets.find((p) => p.id === state.presetId)?.id ?? 'theme';
    downloadBlob(exportTheme(theme), `${name}.json`);
  };

  return (
    <VStack gap={6}>
      <Heading level={2}>Theme</Heading>
      <VStack gap={3}>
        <Selector
          label="Preset"
          hasSearch
          options={presetOptions}
          value={state.presetId ?? IMPORTED_PRESET}
          onChange={(value) => {
            const preset = presets.find((p) => p.id === value);
            if (preset) dispatch({type: 'selectPreset', preset});
          }}
          description="Start from a bundled theme and adjust it below."
        />
        <HStack gap={2} align="end">
          <FileInput
            label="Import theme JSON"
            accept=".json,application/json"
            value={importFile}
            onChange={(file) => void handleImport(file)}
            status={importError ? {type: 'error', message: importError} : undefined}
          />
          <Button label="Export theme JSON" onClick={handleExport} />
        </HStack>
      </VStack>

      <VStack gap={3}>
        <Heading level={3}>Fonts</Heading>
        <Grid columns={3} gap={3}>
          {FONT_SETS.flatMap(({key: set, label, scripts}) =>
            GRID_SCRIPTS.map((script) =>
              scripts.includes(script) ? (
                <TextInput
                  key={fontPath(set, script)}
                  label={`${label} · ${SCRIPT_LABELS[script]}`}
                  value={getFont(theme, set, script)}
                  onChange={(value) => dispatch({type: 'setFont', set, script, value})}
                  isDisabled={set === 'monoFont' && preserveMono}
                  disabledMessage="Unused while the existing monospace fonts are preserved."
                  status={errorStatus(fieldErrors, fontPath(set, script))}
                />
              ) : (
                <div key={fontPath(set, script)} aria-hidden />
              ),
            ),
          )}
        </Grid>
      </VStack>

      <VStack gap={3}>
        <Heading level={3}>Options</Heading>
        <FormLayout>
          <Switch
            label="Bold slide titles"
            value={theme.options.titleBold}
            onChange={(value) => dispatch({type: 'setTitleBold', value})}
          />
          <Selector
            label="First-level bullet weight"
            options={styleOptions}
            value={bodyStyle ?? NO_STYLE}
            onChange={(value) =>
              dispatch({type: 'setBodyFirstLevelStyle', value: value === NO_STYLE ? null : value})
            }
            description="Appended to the body font name for first-level bullets, e.g. “Pretendard SemiBold”. The font must provide the weight as a separate family."
            status={errorStatus(fieldErrors, 'options.bodyFirstLevelStyle')}
          />
          <VStack gap={1}>
            <Switch
              label="Preserve monospace fonts"
              value={preserveMono}
              onChange={(value) => dispatch({type: 'setPreserveMono', value})}
              description="Leave text using a known monospace font untouched instead of replacing it with the monospace fonts above."
            />
            <HStack>
              <MonospaceFontsPopover fonts={monospaceFonts} />
            </HStack>
          </VStack>
        </FormLayout>
        {state.presetId === null && <Text type="supporting">Editing an imported theme.</Text>}
      </VStack>
    </VStack>
  );
}
