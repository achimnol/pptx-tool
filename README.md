# pptx-tool
A script package for pptx helper tools to fill the gap between PowerPoint GUI and OOXML.  
The main function is to fix up pptx font configurations considering Latin/EastAsian/ComplexScript/Symbol typeface mappings.

This script package only relies on `lxml`, a popular XML parsing &amp; manipulation library for Python, and Python 3.10.
It does not require the PowerPoint app to be installed and can be run on any operating system where Python and poetry runs.

## Getting Started

You first need to have [`poetry`](https://python-poetry.org/).
One of the recommended way to install it is to use [`pipx`](https://pypa.github.io/pipx/).
To install Python 3.10, we recommend to use [`pyenv`](https://github.com/pyenv/pyenv).
If you are new to `pyenv` and on macOS, install it using [`homebrew`](https://brew.sh/).

```console
$ poetry install
$ poetry run pptx-tool fix-font --theme=themes/pretendard.json input.pptx output.pptx
```

Check out the `themes` directory for more theme definitions.

By default, the text objects using a known monospace font (see `known_monospace_fonts`
in `pptx_tool/fix.py`) are replaced with the theme's `monoFont`.
If your slides use monospace fonts deliberately, such as for sample code, you may keep them
as-is with the `--preserve-mono` option:

```console
$ poetry run pptx-tool fix-font --preserve-mono --theme=themes/pretendard.json input.pptx output.pptx
```

You may also enable it in the theme file itself, and override it back with `--no-preserve-mono`:

```json
{
  "options": {
    "preserveMono": true
  }
}
```

Note that the preservation works per text object: if any of its typefaces is a monospace font,
the whole object is left untouched, including its symbol typeface and script-specific overrides.
This is intended so that a code block with Korean comments does not get half-converted.

You may also generate and register an office font theme (shared by all Office apps) with
the following command:

```console
$ poetry run pptx-tool generate-font-theme --theme=themes/pretendard.json 'My Pretendard'
```

After restarting the PowerPoint app, you can choose this theme from the "Design" ribbon.

## Known Issues

* After applying the font theme by this tool, there may be multiple major/minor fonts displayed in the font selection list.
  This seems to be due to having all 'latin', 'ea', 'cs' fonts.
  But this is required to prevent issues like incomplete ongoing composition of Hangul characters confusing the PowerPoint app to
  distinguish the correct font to use.

  To keep consistency on new shape objects, it is best to use the "copy style" function to make the fonts consistent and avoid
  using multiple different fonts in the slides.

## Development

```console
$ poetry install
$ poetry run pytest
```
