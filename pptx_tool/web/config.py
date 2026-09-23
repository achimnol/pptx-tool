import dataclasses
from pathlib import Path

from ..package import ArchiveLimits

STATIC_DIR = Path(__file__).parent / "static"
LOOPBACK_ADDRESSES = ("127.0.0.1", "localhost", "::1")


def loopback_host_headers(port: int) -> tuple[str, ...]:
    """The Host header values that address the loopback interface on the given port."""
    hosts = ("127.0.0.1", "localhost", "[::1]")
    return (*hosts, *(f"{host}:{port}" for host in hosts))


@dataclasses.dataclass(frozen=True, kw_only=True)
class WebConfig:
    local: bool = False
    """Allow the endpoints that modify the server machine, such as installing Office font themes."""
    max_upload_size: int = 50 * 1024**2
    archive_limits: ArchiveLimits = dataclasses.field(default_factory=ArchiveLimits)
    static_dir: Path | None = STATIC_DIR
    """The directory containing the built frontend, or None to disable serving it."""
    font_theme_dir: Path | None = None
    """The Office font theme directory, auto-detected when None."""
    allowed_hosts: tuple[str, ...] = loopback_host_headers(8000)
    """The accepted Host header values including the port, which guard against DNS rebinding; empty to accept any."""
