import re
import unicodedata
from datetime import datetime
from pathlib import Path


def build_display_no(task_db_id: int, now: datetime | None = None) -> str:
    dt = now or datetime.now()
    return f"{dt:%Y%m%d}-{task_db_id:06d}"


def sanitize_filename_stem(filename: str | None, max_length: int = 40) -> str:
    stem = Path(filename or "file").stem.strip()
    if not stem:
        return "file"

    ascii_stem = unicodedata.normalize("NFKD", stem).encode("ascii", "ignore").decode("ascii")
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", ascii_stem).strip("_").lower()
    if not cleaned:
        cleaned = re.sub(r"[^0-9A-Za-z]+", "_", stem).strip("_")
    if not cleaned:
        cleaned = "file"
    return cleaned[:max_length]


def ensure_extension(ext: str | None, default_ext: str = ".bin") -> str:
    value = (ext or "").strip()
    if not value:
        return default_ext
    return value if value.startswith(".") else f".{value}"


def build_storage_filename(
    display_no: str,
    original_filename: str | None,
    task_id: str,
    *,
    role: str | None = None,
    ext: str | None = None,
) -> str:
    safe_stem = sanitize_filename_stem(original_filename)
    suffix = ensure_extension(ext or Path(original_filename or "").suffix, ".bin")
    parts = [display_no]
    if role:
        parts.append(role)
    parts.append(safe_stem)
    parts.append(task_id[:8])
    return "_".join(parts) + suffix
