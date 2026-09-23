function stem(filename: string): string {
  const dot = filename.lastIndexOf('.');
  return dot > 0 ? filename.slice(0, dot) : filename;
}

/** The default output file name for a fixed presentation. */
export function defaultOutputName(filename: string): string {
  return `${stem(filename) || 'presentation'}-fixed.pptx`;
}

export function ensurePptxSuffix(filename: string): string {
  const trimmed = filename.trim();
  return trimmed.toLowerCase().endsWith('.pptx') ? trimmed : `${trimmed}.pptx`;
}

export function formatBytes(size: number): string {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KiB`;
  return `${(size / 1024 / 1024).toFixed(1)} MiB`;
}
