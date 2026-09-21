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
in `pptx_tool/fix.py`) have their latin/hangul typefaces replaced with the theme's `monoFont`,
while their symbol typeface is still replaced with the theme's symbol font and their
script-specific overrides are dropped like any other object.
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

Note that `lxml` 4.x has no prebuilt wheels for the recent Python versions, so you may need to
pin the virtualenv to an older interpreter, e.g. `poetry env use 3.11`.
