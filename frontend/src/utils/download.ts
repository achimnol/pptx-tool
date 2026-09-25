/** Let the browser save the resource at the URL as a file with the given name. */
export function downloadUrl(url: string, filename: string): void {
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
}

/** Let the browser save the blob as a file with the given name. */
export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  downloadUrl(url, filename);
  // Revoke in a later task so that the browser can start the download first.
  setTimeout(() => URL.revokeObjectURL(url), 0);
}
