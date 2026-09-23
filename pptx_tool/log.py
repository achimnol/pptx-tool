"""Route the progress messages of the package to the CLI's stdout or to a per-request capture."""

import logging
import sys
import threading
from collections.abc import Iterator
from contextlib import contextmanager
from typing import TextIO

PACKAGE_LOGGER = logging.getLogger("pptx_tool")


def _enable_info_level() -> None:
    # Without an explicit level, the package logger inherits WARNING from the root logger
    # and would drop all progress messages before they reach any handler.
    if PACKAGE_LOGGER.getEffectiveLevel() > logging.INFO:
        PACKAGE_LOGGER.setLevel(logging.INFO)


@contextmanager
def cli_logging(stream: TextIO | None = None) -> Iterator[None]:
    """Print the progress messages as plain lines to stdout (or the given stream) while in the context."""
    handler = logging.StreamHandler(stream or sys.stdout)
    handler.setFormatter(logging.Formatter("%(message)s"))
    _enable_info_level()
    PACKAGE_LOGGER.addHandler(handler)
    try:
        yield
    finally:
        PACKAGE_LOGGER.removeHandler(handler)


class LogCapture:
    """The progress messages captured by `capture_log()`."""

    def __init__(self) -> None:
        self.lines: list[str] = []

    @property
    def text(self) -> str:
        return "".join(f"{line}\n" for line in self.lines)


class _CaptureHandler(logging.Handler):
    def __init__(self, capture: LogCapture, thread_id: int) -> None:
        super().__init__()
        self._capture = capture
        self.setFormatter(logging.Formatter("%(message)s"))
        self.addFilter(lambda record: record.thread == thread_id)

    def emit(self, record: logging.LogRecord) -> None:
        self._capture.lines.append(self.format(record))


@contextmanager
def capture_log() -> Iterator[LogCapture]:
    """
    Capture the progress messages emitted by the current thread while in the context.

    Messages from other threads are ignored, so that concurrent requests handled in
    separate worker threads each get only their own log.
    """
    capture = LogCapture()
    handler = _CaptureHandler(capture, threading.get_ident())
    _enable_info_level()
    PACKAGE_LOGGER.addHandler(handler)
    try:
        yield capture
    finally:
        PACKAGE_LOGGER.removeHandler(handler)
