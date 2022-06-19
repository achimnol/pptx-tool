# pptx-tool
A script package for pptx helper tools.  
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
