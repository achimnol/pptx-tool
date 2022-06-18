import zipfile
from pathlib import Path


def extract_pptx(src_file: Path, dst_dir: Path) -> None:
    with zipfile.ZipFile(src_file) as src:
        src.extractall(dst_dir)


def build_pptx(src_dir: Path, dst_file: Path) -> None:
    dir_queue = []
    with zipfile.ZipFile(dst_file, 'w', compression=zipfile.ZIP_DEFLATED) as dst:
        dir_queue.append(src_dir)
        while dir_queue:
            current_dir = dir_queue.pop()
            for p in current_dir.iterdir():
                if p.is_dir():
                    dir_queue.append(p)
                else:
                    dst.write(p, arcname=str(p.relative_to(src_dir)))
