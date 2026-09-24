import {describe, expect, it} from 'vitest';

import {defaultOutputName, ensurePptxSuffix, formatBytes} from './filenames';

describe('filenames', () => {
  it('derives the default output name', () => {
    expect(defaultOutputName('deck.pptx')).toBe('deck-fixed.pptx');
    expect(defaultOutputName('my.deck.pptx')).toBe('my.deck-fixed.pptx');
    expect(defaultOutputName('.pptx')).toBe('.pptx-fixed.pptx');
    expect(defaultOutputName('')).toBe('presentation-fixed.pptx');
  });

  it('ensures the pptx suffix', () => {
    expect(ensurePptxSuffix(' out ')).toBe('out.pptx');
    expect(ensurePptxSuffix('out.PPTX')).toBe('out.PPTX');
  });

  it('formats sizes', () => {
    expect(formatBytes(10)).toBe('10 B');
    expect(formatBytes(2048)).toBe('2.0 KiB');
    expect(formatBytes(3 * 1024 * 1024)).toBe('3.0 MiB');
  });
});
