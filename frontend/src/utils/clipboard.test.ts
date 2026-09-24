import {describe, expect, it, vi} from 'vitest';

import {copyText} from './clipboard';

function mockExecCommand(result: boolean) {
  const copied: string[] = [];
  const execCommand = vi.fn((command: string) => {
    if (command === 'copy') copied.push((document.activeElement as HTMLTextAreaElement).value);
    return result;
  });
  Object.defineProperty(document, 'execCommand', {value: execCommand, configurable: true});
  return {execCommand, copied};
}

describe('copyText', () => {
  it('writes to the async clipboard API', async () => {
    const writeText = vi.fn(() => Promise.resolve());
    vi.spyOn(navigator, 'clipboard', 'get').mockReturnValue({writeText} as unknown as Clipboard);
    const {execCommand} = mockExecCommand(true);
    await copyText('a path');
    expect(writeText).toHaveBeenCalledWith('a path');
    expect(execCommand).not.toHaveBeenCalled();
  });

  it('falls back to the copy command without the async clipboard API', async () => {
    // Browsers omit navigator.clipboard outside secure contexts, such as a plain-HTTP LAN address.
    vi.spyOn(navigator, 'clipboard', 'get').mockReturnValue(undefined as unknown as Clipboard);
    const {copied} = mockExecCommand(true);
    await copyText('a path');
    expect(copied).toEqual(['a path']);
    expect(document.querySelector('textarea')).toBeNull();
  });

  it('falls back to the copy command when the async clipboard API rejects', async () => {
    const writeText = vi.fn(() => Promise.reject(new DOMException('Denied', 'NotAllowedError')));
    vi.spyOn(navigator, 'clipboard', 'get').mockReturnValue({writeText} as unknown as Clipboard);
    const {copied} = mockExecCommand(true);
    await copyText('a path');
    expect(copied).toEqual(['a path']);
  });

  it('rejects when every method fails', async () => {
    vi.spyOn(navigator, 'clipboard', 'get').mockReturnValue(undefined as unknown as Clipboard);
    mockExecCommand(false);
    await expect(copyText('a path')).rejects.toThrow();
    expect(document.querySelector('textarea')).toBeNull();
  });
});
