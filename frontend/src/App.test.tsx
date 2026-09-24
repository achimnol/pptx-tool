import {render, screen, waitFor} from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import {describe, expect, it, vi} from 'vitest';

import {App} from './App';
import {baseRoutes, mockFetch, THEME, type MockRoute} from './test/fixtures';

function renderApp(routes: MockRoute[] = []) {
  const {fetchMock, requests} = mockFetch([...baseRoutes(), ...routes]);
  vi.stubGlobal('fetch', fetchMock);
  vi.stubGlobal(
    'URL',
    Object.assign(URL, {createObjectURL: vi.fn(() => 'blob:x'), revokeObjectURL: vi.fn()}),
  );
  render(<App />);
  return {requests};
}

function getFileInput(accept: string): HTMLInputElement {
  // FileInput labels both its hidden native input and its trigger button.
  return document.querySelector(`input[type="file"][accept^="${accept}"]`)!;
}

async function waitForEditor() {
  return await screen.findByRole('textbox', {name: /Heading · Latin/});
}

describe('App', () => {
  it('fills the theme editor with the default preset', async () => {
    renderApp();
    expect(await waitForEditor()).toHaveValue('Major Sans');
    expect(screen.getByRole('textbox', {name: /Monospace · Hangul/})).toHaveValue('모노');
  });

  it('places the theme editor before the operations panel', async () => {
    renderApp();
    await waitForEditor();
    const theme = screen.getByRole('heading', {name: 'Theme'});
    const operations = screen.getByRole('heading', {name: 'Operations'});
    expect(
      theme.compareDocumentPosition(operations) & Node.DOCUMENT_POSITION_FOLLOWING,
    ).toBeTruthy();
    // The tabs only switch the operations panel, so they live in the same column.
    const column = operations.closest('.app-columns > *');
    expect(theme.closest('.app-columns > *')?.previousElementSibling).toBeNull();
    expect(column).toContainElement(screen.getByRole('tablist'));
    expect(document.querySelector('.app-content')).toContainElement(theme);
  });

  it('copies the Office theme fonts path', async () => {
    const writeText = vi.fn(() => Promise.resolve());
    vi.spyOn(navigator, 'clipboard', 'get').mockReturnValue({writeText} as unknown as Clipboard);
    renderApp();
    await waitForEditor();
    await userEvent.click(screen.getByRole('tab', {name: 'Export office font theme'}));
    await userEvent.click(screen.getByRole('button', {name: 'Copy the Windows path'}));
    expect(writeText).toHaveBeenCalledWith(
      '%APPDATA%\\Microsoft\\Templates\\Document Themes\\Theme Fonts',
    );
    expect(
      await screen.findByRole('button', {name: 'Copied the Windows path'}),
    ).toBeInTheDocument();
    // The checkmark reverts to the copy icon shortly afterwards.
    expect(
      await screen.findByRole('button', {name: 'Copy the Windows path'}, {timeout: 3000}),
    ).toBeInTheDocument();
  });

  it('copies the Office theme fonts path without the async clipboard API', async () => {
    vi.spyOn(navigator, 'clipboard', 'get').mockReturnValue(undefined as unknown as Clipboard);
    const copied: string[] = [];
    Object.defineProperty(document, 'execCommand', {
      value: vi.fn((command: string) => {
        if (command === 'copy') copied.push((document.activeElement as HTMLTextAreaElement).value);
        return true;
      }),
      configurable: true,
    });
    renderApp();
    await waitForEditor();
    await userEvent.click(screen.getByRole('tab', {name: 'Export office font theme'}));
    await userEvent.click(screen.getByRole('button', {name: 'Copy the macOS path'}));
    expect(copied).toEqual([
      '~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Themes.localized/Theme Fonts',
    ]);
    expect(await screen.findByRole('button', {name: 'Copied the macOS path'})).toBeInTheDocument();
  });

  it('marks an edited preset as modified', async () => {
    renderApp();
    const input = await waitForEditor();
    await userEvent.type(input, ' X');
    expect(input).toHaveValue('Major Sans X');
    expect(screen.getAllByText('Pretendard (modified)')[0]).toBeInTheDocument();
  });

  it('disables the monospace fonts while preserving them', async () => {
    renderApp();
    await waitForEditor();
    const mono = screen.getByRole('textbox', {name: /Monospace · Latin/});
    expect(mono).toBeEnabled();
    await userEvent.click(screen.getByRole('switch', {name: /Preserve monospace fonts/}));
    expect(mono).toHaveAttribute('aria-disabled', 'true');
  });

  it('flags empty font names', async () => {
    renderApp();
    const input = await waitForEditor();
    await userEvent.clear(input);
    await waitFor(() => expect(input).toHaveAttribute('aria-invalid', 'true'));
    expect(screen.getByRole('button', {name: 'Fix fonts'})).toBeDisabled();
  });

  it('fixes the fonts of a presentation', async () => {
    const {requests} = renderApp([
      {
        method: 'POST',
        path: '/api/fix-font',
        body: {
          filename: 'deck-fixed.pptx',
          log: 'Current font scheme\n',
          contentBase64: btoa('PK'),
        },
      },
    ]);
    await waitForEditor();
    const file = new File(['PK'], 'deck.pptx');
    await userEvent.upload(getFileInput('.pptx'), file);
    expect(screen.getByRole('textbox', {name: /Output file name/})).toHaveValue('deck-fixed.pptx');
    await userEvent.click(screen.getByRole('button', {name: 'Fix fonts'}));
    await waitFor(() => expect(URL.createObjectURL).toHaveBeenCalled());
    const request = requests.find((r) => r.path === '/api/fix-font')!;
    const form = request.body as FormData;
    expect((form.get('file') as File).name).toBe('deck.pptx');
    expect(JSON.parse(form.get('theme') as string)).toEqual(THEME);
    expect(await screen.findByText('Processing log')).toBeInTheDocument();
  });

  it('shows the server errors', async () => {
    renderApp([
      {
        method: 'POST',
        path: '/api/fix-font',
        status: 400,
        body: {
          status_code: 400,
          detail: 'Not a valid pptx file: File is not a zip file',
        },
      },
    ]);
    await waitForEditor();
    await userEvent.upload(getFileInput('.pptx'), new File(['x'], 'deck.pptx'));
    await userEvent.click(screen.getByRole('button', {name: 'Fix fonts'}));
    expect(await screen.findByText('Failed to fix the fonts')).toBeInTheDocument();
    expect(screen.getByText('Not a valid pptx file: File is not a zip file')).toBeInTheDocument();
  });

  it('reports a server connection failure', async () => {
    vi.stubGlobal('fetch', () => Promise.reject(new TypeError('Failed to fetch')));
    render(<App />);
    expect(await screen.findByText('Failed to connect to the server')).toBeInTheDocument();
  });
});

describe('App tabs', () => {
  it('keeps the edited theme when switching tabs', async () => {
    renderApp();
    const input = await waitForEditor();
    await userEvent.type(input, ' X');
    await userEvent.click(screen.getByRole('tab', {name: 'Export office font theme'}));
    expect(await screen.findByRole('textbox', {name: /Font theme name/})).toBeInTheDocument();
    expect(screen.getByRole('textbox', {name: /Heading · Latin/})).toHaveValue('Major Sans X');
  });

  it('keeps the chosen presentation when switching tabs', async () => {
    renderApp();
    await waitForEditor();
    await userEvent.upload(getFileInput('.pptx'), new File(['PK'], 'deck.pptx'));
    await userEvent.click(screen.getByRole('tab', {name: 'Export office font theme'}));
    await userEvent.click(screen.getByRole('tab', {name: 'Fix fonts'}));
    expect(screen.getByRole('textbox', {name: /Output file name/})).toHaveValue('deck-fixed.pptx');
  });

  it('imports and exports theme JSON files', async () => {
    renderApp();
    await waitForEditor();
    const imported = {...THEME, majorFont: {...THEME.majorFont, latin: 'Imported Sans'}};
    await userEvent.upload(
      getFileInput('.json'),
      new File([JSON.stringify(imported)], 'custom.json', {type: 'application/json'}),
    );
    expect(await screen.findByRole('textbox', {name: /Heading · Latin/})).toHaveValue(
      'Imported Sans',
    );
    expect(screen.getAllByText('Imported theme')[0]).toBeInTheDocument();

    await userEvent.click(screen.getByRole('button', {name: 'Export theme JSON'}));
    const blob = vi.mocked(URL.createObjectURL).mock.calls.at(-1)![0] as Blob;
    expect(JSON.parse(await blob.text())).toEqual(imported);
  });

  it('reports invalid theme JSON files', async () => {
    renderApp();
    await waitForEditor();
    await userEvent.upload(
      getFileInput('.json'),
      new File(['{"majorFont": {}}'], 'broken.json', {type: 'application/json'}),
    );
    expect(
      (await screen.findAllByText(/majorFont.latin: must be a non-empty string/))[0],
    ).toBeInTheDocument();
    expect(screen.getByRole('textbox', {name: /Heading · Latin/})).toHaveValue('Major Sans');
  });
});
