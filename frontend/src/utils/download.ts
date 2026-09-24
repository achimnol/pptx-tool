/** Let the browser save the blob as a file with the given name. */
export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  // Revoke in a later task so that the browser can start the download first.
  setTimeout(() => URL.revokeObjectURL(url), 0);
}
