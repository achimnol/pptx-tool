/**
 * Copies `text` to the clipboard.
 *
 * Browsers omit the async clipboard API outside secure contexts (such as a plain-HTTP LAN address)
 * and reject it when the clipboard permission is denied, so fall back to the legacy copy command.
 * Rejects when both methods fail.
 */
export async function copyText(text: string): Promise<void> {
  try {
    await navigator.clipboard.writeText(text);
    return;
  } catch {
    // Fall through to the copy command.
  }
  if (!copyWithCommand(text)) throw new Error('Failed to copy to the clipboard.');
}

function copyWithCommand(text: string): boolean {
  const previousFocus = document.activeElement;
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', '');
  // Keep the textarea out of sight and out of the layout while it holds the selection.
  textarea.style.position = 'fixed';
  textarea.style.top = '0';
  textarea.style.left = '0';
  textarea.style.opacity = '0';
  textarea.style.pointerEvents = 'none';
  document.body.append(textarea);
  try {
    textarea.focus();
    textarea.select();
    return document.execCommand('copy');
  } catch {
    return false;
  } finally {
    textarea.remove();
    if (previousFocus instanceof HTMLElement) previousFocus.focus();
  }
}
