import re
import unicodedata
from datetime import datetime
from pathlib import Path

_WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    "COM1",
    "COM2",
    "COM3",
    "COM4",
    "COM5",
    "COM6",
    "COM7",
    "COM8",
    "COM9",
    "LPT1",
    "LPT2",
    "LPT3",
    "LPT4",
    "LPT5",
    "LPT6",
    "LPT7",
    "LPT8",
    "LPT9",
}


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


def sanitize_user_visible_stem(filename: str | None, max_length: int = 80) -> str:
    stem = Path(filename or "file").stem.strip()
    if not stem:
        stem = "file"

    normalized = unicodedata.normalize("NFKC", stem)
    normalized = re.sub(r'[\x00-\x1f<>:"/\\|?*]+', "_", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    normalized = re.sub(r"_+", "_", normalized)
    normalized = normalized.strip(" ._")
    if not normalized:
        normalized = "file"

    if normalized.upper() in _WINDOWS_RESERVED_NAMES:
        normalized = f"{normalized}_file"

    if len(normalized) > max_length:
        normalized = normalized[:max_length].rstrip(" ._")
    return normalized or "file"


def build_user_visible_filename(
    original_filename: str | None,
    *,
    suffix: str | None = None,
    ext: str | None = None,
    max_length: int = 120,
) -> str:
    cleaned_suffix = (suffix or "").strip()
    cleaned_suffix = re.sub(r'[\x00-\x1f<>:"/\\|?*]+', "_", cleaned_suffix)
    cleaned_suffix = re.sub(r"\s+", "_", cleaned_suffix)
    cleaned_suffix = re.sub(r"_+", "_", cleaned_suffix).strip(" ._")

    final_ext = ensure_extension(ext or Path(original_filename or "").suffix, ".bin")
    reserved_length = len(final_ext)
    if cleaned_suffix:
        reserved_length += len(cleaned_suffix) + 1

    stem_max_length = max(16, max_length - reserved_length)
    stem = sanitize_user_visible_stem(original_filename, max_length=stem_max_length)
    final_stem = f"{stem}_{cleaned_suffix}" if cleaned_suffix else stem
    return f"{final_stem}{final_ext}"


def ensure_unique_path(path: Path, existing_path: Path | None = None) -> Path:
    if existing_path is not None:
        try:
            if path.resolve() == existing_path.resolve():
                return path
        except Exception:
            if str(path) == str(existing_path):
                return path

    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    counter = 2
    while True:
        candidate = path.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


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
