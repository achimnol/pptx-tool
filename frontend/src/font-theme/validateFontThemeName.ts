const INVALID_CHARS = /[/\\:*?"<>|]/;

function hasControlChars(value: string): boolean {
  return [...value].some((c) => c.charCodeAt(0) < 0x20 || c.charCodeAt(0) === 0x7f);
}

/**
 * Check that the name is usable as an Office font theme file name, with the same rules as the
 * server (`pptx_tool.fix.validate_font_theme_name()`). Returns an error message or null.
 */
export function validateFontThemeName(name: string): string | null {
  const trimmed = name.trim();
  if (trimmed === '') return 'The font theme name must not be empty.';
  // Count code points like Python's len(), not UTF-16 code units.
  if ([...trimmed].length > 100) return 'The font theme name must be at most 100 characters long.';
  if (INVALID_CHARS.test(trimmed) || hasControlChars(trimmed)) {
    return 'The font theme name must not contain control characters or any of / \\ : * ? " < > |.';
  }
  if (trimmed.endsWith('.')) return 'The font theme name must not end with a dot.';
  return null;
}
