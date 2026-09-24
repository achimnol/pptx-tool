import dataclasses
from collections.abc import Iterator
from pathlib import Path
from typing import Any, Protocol

import pytest

pytest.importorskip("litestar")

from litestar import Litestar
from litestar.testing import TestClient

from pptx_tool.web.app import create_app
from pptx_tool.web.config import WebConfig


class MakeClient(Protocol):
    def __call__(self, **overrides: Any) -> TestClient[Litestar]: ...


@pytest.fixture
def font_theme_dir(tmp_path: Path) -> Path:
    path = tmp_path / "Theme Fonts"
    path.mkdir()
    return path


@pytest.fixture
def tmp_dir(tmp_path: Path) -> Path:
    """The temporary storage root of the server under test, kept away from the real /tmp/pptx-tool."""
    return tmp_path / "pptx-tool"


@pytest.fixture
def make_client(font_theme_dir: Path, tmp_dir: Path) -> Iterator[MakeClient]:
    clients: list[TestClient[Litestar]] = []

    def _make_client(**overrides: Any) -> TestClient[Litestar]:
        web_config = dataclasses.replace(
            WebConfig(static_dir=None, font_theme_dir=font_theme_dir, allowed_hosts=(), tmp_dir=tmp_dir),
            **overrides,
        )
        client = TestClient(create_app(web_config))
        client.__enter__()
        clients.append(client)
        return client

    yield _make_client
    for client in clients:
        client.__exit__(None, None, None)


@pytest.fixture
def client(make_client: MakeClient) -> TestClient[Litestar]:
    return make_client()
