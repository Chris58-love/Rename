import re
import shutil
from datetime import datetime
from pathlib import Path

from .config import MAX_STEM_LENGTH, WINDOWS_RESERVED_NAMES


def is_colab() -> bool:
    try:
        import google.colab  # type: ignore  # noqa: F401
        return True
    except Exception:
        return False


def ensure_dir(path) -> Path:
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    return path


def safe_remove(path) -> None:
    path = Path(path)
    if not path.exists():
        return
    if path.is_dir():
        shutil.rmtree(path)
    else:
        path.unlink()


def timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def normalize_text(value) -> str:
    return re.sub(r"\s+", " ", "" if value is None else str(value)).strip()


def parse_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "y", "是", "对", "保留"}


def sanitize_filename_component(name: str) -> str:
    name = "" if name is None else str(name)
    name = re.sub(r"[\x00-\x1f\x7f]", "", name)
    name = re.sub(r'[\\/:\*\?"<>\|]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = name.strip(" .")
    if not name:
        name = "untitled"
    if name.upper() in WINDOWS_RESERVED_NAMES:
        name += "_"
    name = name[:MAX_STEM_LENGTH].strip(" .")
    return name or "untitled"


def split_filename(filename: str):
    path = Path(filename)
    suffix = path.suffix
    if suffix and path.name != suffix:
        return path.name[:-len(suffix)], suffix, True
    return path.name, "", False


def strip_same_extension(pasted_name: str, original_extension: str) -> str:
    value = normalize_text(pasted_name)
    if original_extension and value.lower().endswith(original_extension.lower()):
        value = value[:-len(original_extension)]
    return sanitize_filename_component(value)


def prefix_number(name: str) -> str:
    match = re.match(r"^\s*(\d+)", str(name or ""))
    return match.group(1) if match else ""
