from pathlib import Path

from .utils import split_filename


def build_file_inventory(paths):
    inventory = []
    for raw_path in paths:
        path = Path(raw_path)
        stem, extension, has_extension = split_filename(path.name)
        inventory.append({
            "original_name": path.name,
            "stem": stem,
            "extension": extension,
            "has_extension": has_extension,
            "source_path": str(path.resolve()),
            "size_bytes": path.stat().st_size if path.exists() else 0,
        })
    return inventory
