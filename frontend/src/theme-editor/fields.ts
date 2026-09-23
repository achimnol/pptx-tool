import type {ThemeData} from '../api/types';

export type FontSetKey = 'majorFont' | 'minorFont' | 'monoFont';
export type ScriptKey = 'latin' | 'hangul' | 'symbol';

export interface FontSetSpec {
  key: FontSetKey;
  label: string;
  scripts: readonly ScriptKey[];
}

export const FONT_SETS: readonly FontSetSpec[] = [
  {key: 'majorFont', label: 'Heading', scripts: ['latin', 'hangul', 'symbol']},
  {key: 'minorFont', label: 'Body', scripts: ['latin', 'hangul', 'symbol']},
  {key: 'monoFont', label: 'Monospace', scripts: ['latin', 'hangul']},
];

export const SCRIPT_LABELS: Record<ScriptKey, string> = {
  latin: 'Latin',
  hangul: 'Hangul',
  symbol: 'Symbol',
};

/** The weight suffixes commonly available as separate font families. */
export const BODY_FIRST_LEVEL_STYLES = ['Medium', 'SemiBold', 'Bold', 'ExtraBold'] as const;

export function getFont(theme: ThemeData, set: FontSetKey, script: ScriptKey): string {
  const fontSet = theme[set] as Partial<Record<ScriptKey, string>>;
  return fontSet[script] ?? '';
}

export function fontPath(set: FontSetKey, script: ScriptKey): string {
  return `${set}.${script}`;
}
