import type {components} from './schema';

type Schemas = components['schemas'];

export type ThemeData = Schemas['ThemeData'];
export type BundledThemeInfo = Schemas['BundledThemeInfo'];
export type AppConfig = Schemas['AppConfig'];

/** The parts of the multipart/form-data response of the fix-font request. */
export interface FixFontResult extends Omit<Schemas['FixFontResult'], 'file'> {
  file: File;
}
