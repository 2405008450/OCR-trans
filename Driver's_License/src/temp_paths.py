from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any


MODULE_ROOT = Path(__file__).resolve().parents[1]
CACHE_ROOT = MODULE_ROOT / ".cache"
RUNTIME_TEMP_ROOT = CACHE_ROOT / "tmp"
TEST_TEMP_ROOT = CACHE_ROOT / "test_tmp"


def _ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_runtime_temp_dir() -> Path:
    return _ensure_dir(RUNTIME_TEMP_ROOT)


def get_test_temp_dir() -> Path:
    return _ensure_dir(TEST_TEMP_ROOT)


def create_named_temporary_file(
    *,
    mode: str = "w+b",
    suffix: str = "",
    prefix: str = "drivers_license_",
    delete: bool = False,
    dir: str | Path | None = None,
    **kwargs: Any,
):
    temp_dir = _ensure_dir(Path(dir) if dir is not None else get_runtime_temp_dir())
    return tempfile.NamedTemporaryFile(
        mode=mode,
        suffix=suffix,
        prefix=prefix,
        delete=delete,
        dir=str(temp_dir),
        **kwargs,
    )


def create_temp_dir(
    *,
    prefix: str = "drivers_license_",
    dir: str | Path | None = None,
) -> Path:
    temp_dir = _ensure_dir(Path(dir) if dir is not None else get_test_temp_dir())
    return Path(tempfile.mkdtemp(prefix=prefix, dir=str(temp_dir)))
