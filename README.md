# pptx-font-fix
A script to fix up pptx font configurations considering Latin/EastAsian/ComplexScript/Symbol typeface mappings

This script package only relies on `lxml`, a popular XML parsing &amp; manipulation library for Python, and Python 3.10.
It does not require the PowerPoint app to be installed and can be run on any operating system where Python and poetry runs.

## Getting Started

You first need to have [`poetry`](https://python-poetry.org/).
One of the recommended way to install it is to use [`pipx`](https://pypa.github.io/pipx/).
To install Python 3.10, we recommend to use [`pyenv`](https://github.com/pyenv/pyenv).
If you are new to `pyenv` and on macOS, install it using [`homebrew`](https://brew.sh/).

```console
$ poetry install
$ poetry run pptx-tool fix-font --theme=themes/pretendard.json pptx-font-fix input.pptx output.pptx
```

Check out the `themes` directory for more theme definitions.

You may also generate and register an office font theme (shared by all Office apps) with
the following command:

```console
$ poetry run pptx-tool generate-font-theme --theme=themes/pretendard.json 'My Pretendard'
```

After restarting the PowerPoint app, you can choose this theme from the "Design" ribbon.