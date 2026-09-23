import {render, screen, waitFor} from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import {describe, expect, it, vi} from 'vitest';

import {mockFetch, THEME, type MockRoute} from '../test/fixtures';
import {FontThemePanel} from './FontThemePanel';

function renderPanel(routes: MockRoute[], isLocal: boolean) {
  const {fetchMock, requests} = mockFetch(routes);
  vi.stubGlobal('fetch', fetchMock);
  URL.createObjectURL = vi.fn(() => 'blob:x');
  URL.revokeObjectURL = vi.fn();
  render(<FontThemePanel theme={THEME} isThemeValid isLocal={isLocal} onFieldErrors={() => {}} />);
  return {requests};
}

const installBody = (overwrite: boolean) => ({name: 'My Theme', theme: THEME, overwrite});

describe('FontThemePanel', () => {
  it('hides the install button unless running locally', () => {
    renderPanel([], false);
    expect(screen.queryByRole('button', {name: 'Install to Office'})).not.toBeInTheDocument();
  });

  it('validates the name', async () => {
    renderPanel([], false);
    const download = screen.getByRole('button', {name: 'Download XML'});
    expect(download).toBeDisabled();
    await userEvent.type(screen.getByRole('textbox', {name: /Font theme name/}), 'a/b');
    expect(download).toBeDisabled();
    expect(screen.getByRole('textbox', {name: /Font theme name/})).toHaveAttribute(
      'aria-invalid',
      'true',
    );
  });

  it('downloads the font theme XML', async () => {
    const {requests} = renderPanel(
      [{method: 'POST', path: '/api/font-theme', body: '<a:fontScheme/>', isBlob: true}],
      false,
    );
    await userEvent.type(screen.getByRole('textbox', {name: /Font theme name/}), 'My Theme');
    await userEvent.click(screen.getByRole('button', {name: 'Download XML'}));
    await waitFor(() => expect(URL.createObjectURL).toHaveBeenCalled());
    expect(JSON.parse(requests[0]!.body as string)).toEqual({name: 'My Theme', theme: THEME});
  });

  it('asks before overwriting an installed font theme', async () => {
    const {requests} = renderPanel(
      [
        {
          method: 'POST',
          path: '/api/font-theme/install',
          status: 409,
          body: {status_code: 409, detail: "A font theme named 'My Theme' already exists."},
        },
        {
          method: 'POST',
          path: '/api/font-theme/install',
          status: 201,
          body: {path: '/x/My Theme.xml'},
        },
      ],
      true,
    );
    await userEvent.type(screen.getByRole('textbox', {name: /Font theme name/}), 'My Theme');
    await userEvent.click(screen.getByRole('button', {name: 'Install to Office'}));
    const overwrite = await screen.findByRole('button', {name: 'Overwrite'});
    expect(JSON.parse(requests[0]!.body as string)).toEqual(installBody(false));
    await userEvent.click(overwrite);
    await waitFor(() => expect(requests).toHaveLength(2));
    expect(JSON.parse(requests[1]!.body as string)).toEqual(installBody(true));
    // The toast is also announced through a live region.
    expect(
      await screen.findAllByText(/Installed the font theme at \/x\/My Theme.xml/),
    ).not.toHaveLength(0);
  });

  it('shows other install errors', async () => {
    renderPanel(
      [
        {
          method: 'POST',
          path: '/api/font-theme/install',
          status: 503,
          body: {status_code: 503, detail: 'The office theme directory does not exist.'},
        },
      ],
      true,
    );
    await userEvent.type(screen.getByRole('textbox', {name: /Font theme name/}), 'My Theme');
    await userEvent.click(screen.getByRole('button', {name: 'Install to Office'}));
    expect(
      await screen.findByText('The office theme directory does not exist.'),
    ).toBeInTheDocument();
  });
});
