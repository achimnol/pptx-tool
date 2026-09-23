from pathlib import Path

from .conftest import MakeClient


def test_serves_built_frontend(make_client: MakeClient, tmp_path: Path) -> None:
    static_dir = tmp_path / "static"
    (static_dir / "assets").mkdir(parents=True)
    (static_dir / "index.html").write_text("<html>frontend</html>")
    (static_dir / "assets" / "app.js").write_text("console.log('app')")
    client = make_client(static_dir=static_dir)
    assert client.get("/").text == "<html>frontend</html>"
    assert client.get("/assets/app.js").text == "console.log('app')"
    assert client.get("/api/monospace-fonts").status_code == 200
    assert client.get("/schema/openapi.json").status_code == 200


def test_frontend_not_built(make_client: MakeClient, tmp_path: Path) -> None:
    client = make_client(static_dir=tmp_path / "missing")
    resp = client.get("/")
    assert resp.status_code == 200
    assert "not built" in resp.text
