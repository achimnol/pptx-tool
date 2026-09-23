import dataclasses
import zipfile
from pathlib import Path
from typing import IO


@dataclasses.dataclass(frozen=True)
class ArchiveLimits:
    """Limits applied when extracting untrusted archives, to reject zip bombs."""

    max_entries: int = 20_000
    max_total_size: int = 2 * 1024**3
    """The maximum total uncompressed size in bytes."""


class UnsafeArchiveError(ValueError):
    """Raised when an archive exceeds the given extraction limits."""


def extract_pptx(src_file: Path | IO[bytes], dst_dir: Path, *, limits: ArchiveLimits | None = None) -> None:
    with zipfile.ZipFile(src_file) as src:
        if limits is not None:
            # The declared sizes are enforced while extracting: a member producing more data
            # than its header claims fails with a CRC error.
            members = src.infolist()
            if len(members) > limits.max_entries:
                raise UnsafeArchiveError(f"The archive has too many entries (more than {limits.max_entries}).")
            if sum(m.file_size for m in members) > limits.max_total_size:
                raise UnsafeArchiveError(
                    f"The archive is too large when uncompressed (more than {limits.max_total_size} bytes)."
                )
        src.extractall(dst_dir)


def build_pptx(src_dir: Path, dst_file: Path) -> None:
    dir_queue: list[Path] = []
    with zipfile.ZipFile(dst_file, "w", compression=zipfile.ZIP_DEFLATED) as dst:
        dir_queue.append(src_dir)
        while dir_queue:
            current_dir = dir_queue.pop()
            for p in current_dir.iterdir():
                if p.is_dir():
                    dir_queue.append(p)
                else:
                    if p.parent.name == "themes" and p.name != "theme1.xml":
                        continue
                    dst.write(p, arcname=str(p.relative_to(src_dir)))
