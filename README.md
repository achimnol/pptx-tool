# pptx-tool
A script package for pptx helper tools to fill the gap between PowerPoint GUI and OOXML.
The main function is to fix up pptx font configurations considering Latin/EastAsian/ComplexScript/Symbol typeface mappings.

This script package only relies on `lxml`, a popular XML parsing &amp; manipulation library for Python, and Python 3.13 or later.
It does not require the PowerPoint app to be installed and can be run on any operating system where Python and uv runs.

## Getting Started

You first need to have [`uv`](https://docs.astral.sh/uv/), which manages both the Python
interpreter and the virtualenv for you.

```console
$ uv sync
$ uv run pptx-tool fix-font --theme=pretendard input.pptx output.pptx
```

The `--theme` option takes either the id of a bundled theme, such as `pretendard`, or the path to
a theme JSON file.  Check out the `pptx_tool/themes` directory for the bundled theme definitions.
The bundled themes, with the display names shown in the web UI (major / minor fonts), are:

- `freesentation` — Freesentation
- `inter-pretendard` — Inter Display + Pretendard
- `nanum-barun` — NanumBarunGothic
- `nanum-square` — NanumSquare ExtraBold / NanumSquare
- `nanum-square-round` — NanumSquareRound ExtraBold / NanumSquareRound
- `office` — Office (Calibri + Malgun Gothic)
- `paperlogy` — Paperlogy
- `paperlogy-freesentation` — Paperlogy / Freesentation
- `paperlogy-snu-appendard` — Paperlogy / SNU Appendard
- `pretendard` — Pretendard
- `snu-appendard` — SNU Appendard

By default, the text objects using a known monospace font (see `known_monospace_fonts`
in `pptx_tool/fix.py`) have their latin/hangul typefaces replaced with the theme's `monoFont`,
while their symbol typeface is still replaced with the theme's symbol font and their
script-specific overrides are dropped like any other object.
If your slides use monospace fonts deliberately, such as for sample code, you may keep them
as-is with the `--preserve-mono` option:

```console
$ uv run pptx-tool fix-font --preserve-mono --theme=pretendard input.pptx output.pptx
```

You may also enable it in the theme file itself, and override it back with `--no-preserve-mono`:

```json
{
  "options": {
    "preserveMono": true
  }
}
```

Note that the preservation works per text property element, that is, a single run's `rPr` or a
paragraph's `defRPr`, rather than per shape: if any of its typefaces is a monospace font, that
whole element is left untouched, including its symbol typeface and script-specific overrides.
This is intended so that a code block with Korean comments does not have only its latin part
converted.  Bullet fonts using a monospace font are preserved likewise.

The matching is an exact, case-insensitive lookup of the known font list, so it does not cover
the objects whose typeface carries a weight suffix such as "JetBrains Mono ExtraBold".
Since the unit is a single run, an empty line of a code block still gets converted if
PowerPoint has stamped a non-monospace typeface on its `endParaRPr`.

You may also generate and register an office font theme (shared by all Office apps) with
the following command:

```console
$ uv run pptx-tool generate-font-theme --theme=pretendard 'My Pretendard'
```

After restarting the PowerPoint app, you can choose this theme from the "Design" ribbon.

## Web UI

`pptx-tool` also provides a web UI to fix up presentations and generate Office font themes from
the browser, editing the theme interactively.  It requires the optional `web` extra and a built
frontend ([Node.js](https://nodejs.org/) 24 and [pnpm](https://pnpm.io/)):

```console
$ uv sync --extra web
$ pnpm -C frontend install
$ pnpm -C frontend build
$ uv run pptx-tool serve
```

Then open http://127.0.0.1:8000 in the browser.  The server only listens on the loopback
interface by default; use `--host` and `--port` to change it, and `--max-upload-mb` to change the
upload size limit (200 MiB by default).
The uploaded and intermediate files are kept in per-request subdirectories of `/tmp/pptx-tool`
(`--tmp-dir`), limited to 1 GiB in total (`--tmp-quota-mb`); each request is deleted when done, and
any leftovers of earlier requests are deleted to make room.  Give each server process its own
directory.  Note that the server framework spools each upload in the system temporary directory
while the request is processed, so the peak disk usage is the quota plus the concurrent uploads.
When listening on a loopback address, the server checks the `Host` header against DNS rebinding,
so access it via `127.0.0.1`, `localhost` or `[::1]` with the same port as the server, not through
a port-forwarding tunnel or a reverse proxy with a different port.

Installing Office font themes from the web UI writes into the Office theme directory of the
machine running the server, so it is available only when the server runs with `--local`, which
is allowed only on a loopback address:

```console
$ uv run pptx-tool serve --local
```

### Container image

The `Dockerfile` builds the web UI image with the `web` extra and the built frontend.
Tagged releases (`v*`) are published to `ghcr.io/achimnol/pptx-tool` by the `Image` workflow,
and the latest main is published as the `edge` tag.
The image runs as a non-root user and needs only `/tmp` to be writable:

```console
$ docker build -t pptx-tool .
$ docker run --rm --read-only --tmpfs /tmp -p 8000:8000 pptx-tool
```

The default command is `serve --host=0.0.0.0 --port=8000`; pass your own arguments to override it,
e.g. `--max-upload-mb=100`.  `--local` is refused on a non-loopback address, so it is not available
in the container.

## Fonts

The bundled themes use the following fonts, which are not shipped with Office.
Install them from their official pages before opening the fixed presentations:

- Pretendard: https://cactus.tistory.com/306
- SNU Appendard: https://qbio.io/share/fonts/
- Inter and Inter Display: https://fonts.google.com/specimen/Inter
- Sarasa Term K: https://picaq.github.io/sarasa
- Paperlogy: https://freesentation.blog/paperlogyfont
- Freesentation: https://freesentation.blog/freesentation
- Nanum Fonts: https://hangeul.naver.com/font

The web UI lists the same links under the *Where to download the fonts?* button of the theme editor.

## Known Issues

* After applying the font theme by this tool, there may be multiple major/minor fonts displayed in the font selection list.
  This seems to be due to having all 'latin', 'ea', 'cs' fonts.
  But this is required to prevent issues like incomplete ongoing composition of Hangul characters confusing the PowerPoint app to
  distinguish the correct font to use.

  To keep consistency on new shape objects, it is best to use the "copy style" function to make the fonts consistent and avoid
  using multiple different fonts in the slides.

## Development

```console
$ uv sync --all-groups --all-extras
$ uv run pytest
```

The repository is linted and formatted with [`ruff`](https://docs.astral.sh/ruff/) and
type-checked with [`mypy`](https://mypy-lang.org/) in strict mode.  Install the git hooks once
so that the same checks run before each commit:

```console
$ uv run pre-commit install
```

You can also run every check manually:

```console
$ uv run ruff check .
$ uv run ruff format .
$ uv run mypy
$ uv run pre-commit run --all-files
```

The same checks run in CI for every push and pull request, and the test suite runs against
Python 3.13 and 3.14.

### Web UI development

Run the backend and the Vite dev server side by side; the dev server proxies the API requests to
the backend:

```console
$ uv run pptx-tool serve
$ pnpm -C frontend dev
```

The frontend is checked with `pnpm -C frontend typecheck`, `pnpm -C frontend lint` and
`pnpm -C frontend test`.  The frontend API types are generated from the OpenAPI schema of the
backend, so regenerate them after changing the API models:

```console
$ pnpm -C frontend gen:openapi
$ pnpm -C frontend gen:api
```
