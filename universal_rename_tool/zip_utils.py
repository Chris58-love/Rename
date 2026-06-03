from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from .utils import ensure_dir


def make_zip(paths, zip_path):
    zip_path = Path(zip_path)
    ensure_dir(zip_path.parent)
    with ZipFile(zip_path, "w", ZIP_DEFLATED) as zf:
        for path in paths:
            path = Path(path)
            if path.exists() and path.is_file():
                zf.write(path, arcname=path.name)
    return zip_path
