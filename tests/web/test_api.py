import base64
import copy
import io
import json
import zipfile
from pathlib import Path
from typing import Any

import pytest
from litestar import Litestar
from litestar.testing import TestClient
from lxml import etree

from pptx_tool.package import ArchiveLimits
from pptx_tool.web.config import loopback_host_headers

from .conftest import MakeClient

PPTX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

THEME: dict[str, Any] = {
    "majorFont": {"latin": "Major Sans", "hangul": "메이저", "symbol": "Major Symbol"},
    "minorFont": {"latin": "Minor Sans", "hangul": "마이너", "symbol": "Minor Symbol"},
    "monoFont": {"latin": "Mono Code", "hangul": "모노"},
    "options": {"titleBold": True, "bodyFirstLevelStyle": "SemiBold", "preserveMono": False},
}


def _post_fix_font(
    client: TestClient[Litestar],
    content: bytes,
    theme: Any = THEME,
    filename: str = "sample.pptx",
) -> Any:
    theme_field = theme if isinstance(theme, str) else json.dumps(theme)
    return client.post(
        "/api/fix-font",
        files={"file": (filename, content, PPTX_MEDIA_TYPE)},
        data={"theme": theme_field},
    )


def test_list_themes(client: TestClient[Litestar]) -> None:
    resp = client.get("/api/themes")
    assert resp.status_code == 200
    themes = resp.json()
    assert len(themes) >= 11
    pretendard = next(t for t in themes if t["id"] == "pretendard")
    assert pretendard["name"] == "Pretendard"
    assert pretendard["theme"]["majorFont"]["latin"] == "Pretendard"
    assert pretendard["theme"]["options"] == {
        "titleBold": True,
        "bodyFirstLevelStyle": "SemiBold",
        "preserveMono": False,
    }


def test_list_monospace_fonts(client: TestClient[Litestar]) -> None:
    fonts = client.get("/api/monospace-fonts").json()
    assert fonts == sorted(fonts)
    assert "consolas" in fonts


def test_get_config(make_client: MakeClient) -> None:
    assert make_client().get("/api/config").json() == {"local": False, "maxUploadSize": 200 * 1024**2}
    assert make_client(local=True, max_upload_size=1024).get("/api/config").json() == {
        "local": True,
        "maxUploadSize": 1024,
    }


def test_fix_font(client: TestClient[Litestar], sample_pptx: Path) -> None:
    resp = _post_fix_font(client, sample_pptx.read_bytes())
    assert resp.status_code == 200, resp.text
    result = resp.json()
    assert result["filename"] == "sample-fixed.pptx"
    assert "Current font scheme: (name='Office')" in result["log"]
    assert "slide1.xml: normal element" in result["log"]
    with zipfile.ZipFile(io.BytesIO(base64.b64decode(result["contentBase64"]))) as zf:
        theme_xml = zf.read("ppt/theme/theme1.xml").decode()
        slide_xml = zf.read("ppt/slides/slide1.xml").decode()
    assert 'typeface="Major Sans"' in theme_xml
    assert 'typeface="+mn-lt"' in slide_xml
    assert 'typeface="Mono Code"' in slide_xml


def test_fix_font_preserve_mono(client: TestClient[Litestar], sample_pptx: Path) -> None:
    theme = copy.deepcopy(THEME)
    theme["options"]["preserveMono"] = True
    result = _post_fix_font(client, sample_pptx.read_bytes(), theme).json()
    with zipfile.ZipFile(io.BytesIO(base64.b64decode(result["contentBase64"]))) as zf:
        slide_xml = zf.read("ppt/slides/slide1.xml").decode()
    assert 'typeface="Consolas"' in slide_xml
    assert "Preserving the existing monospace fonts as-is." in result["log"]


def test_fix_font_invalid_theme_value(client: TestClient[Litestar], sample_pptx: Path) -> None:
    theme = copy.deepcopy(THEME)
    theme["majorFont"]["latin"] = " "
    resp = _post_fix_font(client, sample_pptx.read_bytes(), theme)
    assert resp.status_code == 400
    assert resp.json()["extra"] == [
        {"key": "majorFont.latin", "message": "must be a non-empty string", "source": "body"}
    ]


@pytest.mark.parametrize("theme", ["{", json.dumps({"majorFont": {}}), json.dumps([])])
def test_fix_font_malformed_theme(client: TestClient[Litestar], sample_pptx: Path, theme: str) -> None:
    resp = _post_fix_font(client, sample_pptx.read_bytes(), theme)
    assert resp.status_code == 400
    assert "theme" in resp.json()["detail"].lower()


def test_fix_font_not_a_zip(client: TestClient[Litestar]) -> None:
    resp = _post_fix_font(client, b"not a zip file")
    assert resp.status_code == 400
    assert resp.json()["detail"].startswith("Not a valid pptx file")


def test_fix_font_archive_limits(make_client: MakeClient, sample_pptx: Path) -> None:
    client = make_client(archive_limits=ArchiveLimits(max_entries=3))
    resp = _post_fix_font(client, sample_pptx.read_bytes())
    assert resp.status_code == 400
    assert "too many entries" in resp.json()["detail"]


def test_fix_font_not_a_pptx(client: TestClient[Litestar]) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("hello.txt", "hello")
    resp = _post_fix_font(client, buf.getvalue())
    assert resp.status_code == 422
    assert resp.json()["detail"].startswith("Failed to process the presentation")


def test_fix_font_too_large(make_client: MakeClient, sample_pptx: Path) -> None:
    client = make_client(max_upload_size=1024)
    content = sample_pptx.read_bytes()
    assert len(content) > 1024
    resp = _post_fix_font(client, content)
    assert resp.status_code == 413


def test_font_theme_download(client: TestClient[Litestar]) -> None:
    resp = client.post("/api/font-theme", json={"name": "나의 테마", "theme": THEME})
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("application/xml")
    assert "filename*=UTF-8''%EB%82%98%EC%9D%98%20%ED%85%8C%EB%A7%88.xml" in resp.headers["content-disposition"]
    root_elem = etree.fromstring(resp.content)
    assert root_elem.get("name") == "나의 테마"


@pytest.mark.parametrize("path", ["/api/font-theme", "/api/font-theme/install"])
def test_font_theme_invalid_name(make_client: MakeClient, font_theme_dir: Path, path: str) -> None:
    client = make_client(local=True)
    resp = client.post(path, json={"name": "../evil", "theme": THEME})
    assert resp.status_code == 400
    assert resp.json()["extra"][0]["key"] == "name"
    assert list(font_theme_dir.parent.glob("*.xml")) == []


def test_font_theme_invalid_theme(client: TestClient[Litestar]) -> None:
    theme = copy.deepcopy(THEME)
    theme["minorFont"]["symbol"] = ""
    resp = client.post("/api/font-theme", json={"name": "x", "theme": theme})
    assert resp.status_code == 400
    assert resp.json()["extra"][0]["key"] == "minorFont.symbol"


def test_install_font_theme_requires_local_mode(client: TestClient[Litestar], font_theme_dir: Path) -> None:
    resp = client.post("/api/font-theme/install", json={"name": "My Theme", "theme": THEME})
    assert resp.status_code == 403
    assert not (font_theme_dir / "My Theme.xml").exists()


def test_install_font_theme(make_client: MakeClient, font_theme_dir: Path) -> None:
    client = make_client(local=True)
    body: dict[str, Any] = {"name": "My Theme", "theme": THEME}
    resp = client.post("/api/font-theme/install", json=body)
    assert resp.status_code == 201
    path = font_theme_dir / "My Theme.xml"
    assert resp.json() == {"path": str(path)}
    assert etree.fromstring(path.read_bytes()).get("name") == "My Theme"

    resp = client.post("/api/font-theme/install", json=body)
    assert resp.status_code == 409
    assert resp.json()["extra"] == {"path": str(path)}

    resp = client.post("/api/font-theme/install", json={**body, "overwrite": True})
    assert resp.status_code == 201


def test_install_font_theme_missing_dir(make_client: MakeClient, tmp_path: Path) -> None:
    client = make_client(local=True, font_theme_dir=tmp_path / "missing")
    resp = client.post("/api/font-theme/install", json={"name": "My Theme", "theme": THEME})
    assert resp.status_code == 503
    assert "does not exist" in resp.json()["detail"]


def test_install_font_theme_rejects_cross_origin(make_client: MakeClient, font_theme_dir: Path) -> None:
    client = make_client(local=True)
    body = {"name": "My Theme", "theme": THEME}
    resp = client.post("/api/font-theme/install", json=body, headers={"Origin": "https://evil.example"})
    assert resp.status_code == 403
    assert not (font_theme_dir / "My Theme.xml").exists()
    resp = client.post("/api/font-theme/install", json=body, headers={"Origin": "http://testserver.local"})
    assert resp.status_code == 201


def test_allowed_hosts(make_client: MakeClient) -> None:
    client = make_client(allowed_hosts=loopback_host_headers(8000))
    assert client.get("/api/config").status_code == 400
    assert client.get("/api/config", headers={"Host": "evil.example:8000"}).status_code == 400
    assert client.get("/api/config", headers={"Host": "localhost:8001"}).status_code == 400
    for host in ("localhost:8000", "127.0.0.1:8000", "[::1]:8000", "localhost"):
        assert client.get("/api/config", headers={"Host": host}).status_code == 200


def test_openapi_schema(client: TestClient[Litestar]) -> None:
    schema = client.get("/schema/openapi.json").json()
    theme_data = schema["components"]["schemas"]["ThemeData"]
    assert set(theme_data["properties"]) == {"majorFont", "minorFont", "monoFont", "options"}
    assert "titleBold" in schema["components"]["schemas"]["ThemeOptions"]["properties"]
    assert "/api/font-theme/install" in schema["paths"]
