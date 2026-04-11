import os
import shutil
import subprocess
import threading
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator


LIBREOFFICE_PATH = os.getenv("LIBREOFFICE_PATH", "").strip()
_LIBREOFFICE_LOCK = threading.Lock()


def resolve_libreoffice_path(configured_path: str | None = None) -> str:
    candidates: list[str] = []

    for candidate in [
        configured_path,
        os.getenv("LIBREOFFICE_PATH", "").strip(),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "soffice",
        "libreoffice",
    ]:
        if candidate and candidate not in candidates:
            candidates.append(candidate)

    for candidate in candidates:
        if any(sep in candidate for sep in ("/", "\\")) or Path(candidate).is_absolute():
            if Path(candidate).exists():
                return str(Path(candidate))
            continue

        resolved = shutil.which(candidate)
        if resolved:
            return resolved

    raise FileNotFoundError(
        "未找到 LibreOffice 可执行文件。请安装 LibreOffice，并在环境变量 LIBREOFFICE_PATH "
        "中指定 soffice 路径，或确保 `soffice` 已加入 PATH。"
    )


@contextmanager
def _temporary_profile_dir(base_dir: Path) -> Iterator[Path]:
    profile_dir = base_dir / f".libreoffice-profile-{uuid.uuid4().hex}"
    profile_dir.mkdir(parents=True, exist_ok=True)
    try:
        yield profile_dir
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


def _run_libreoffice_convert(
    input_path: str | Path,
    output_dir: str | Path,
    convert_to: str,
    *,
    libreoffice_path: str | None = None,
) -> subprocess.CompletedProcess[str]:
    input_file = Path(input_path).resolve()
    target_dir = Path(output_dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)

    soffice = resolve_libreoffice_path(libreoffice_path)
    with _temporary_profile_dir(target_dir) as profile_dir:
        command = [
            str(soffice),
            f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
            "--headless",
            "--convert-to",
            convert_to,
            "--outdir",
            str(target_dir),
            str(input_file),
        ]
        with _LIBREOFFICE_LOCK:
            return subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
            )


def convert_to_docx_via_libreoffice(
    input_path: str | Path,
    output_path: str | Path | None = None,
    *,
    libreoffice_path: str | None = None,
) -> str:
    input_file = Path(input_path).resolve()
    if not input_file.exists():
        raise FileNotFoundError(f"待转换文件不存在: {input_file}")

    if output_path is None:
        output_file = input_file.with_suffix(".docx")
    else:
        output_file = Path(output_path).resolve()
    output_file.parent.mkdir(parents=True, exist_ok=True)

    if input_file.suffix.lower() == ".docx":
        if output_file != input_file:
            if output_file.exists():
                output_file.unlink()
            shutil.copy2(input_file, output_file)
        return str(output_file)

    expected_docx = output_file.parent / f"{input_file.stem}.docx"
    if expected_docx.exists() and expected_docx != input_file:
        expected_docx.unlink()
    if output_file.exists() and output_file != input_file and output_file != expected_docx:
        output_file.unlink()

    result = _run_libreoffice_convert(
        input_path=input_file,
        output_dir=output_file.parent,
        convert_to="docx:Office Open XML Text",
        libreoffice_path=libreoffice_path,
    )
    if result.returncode != 0:
        raise RuntimeError(
            "LibreOffice 转换失败: "
            f"returncode={result.returncode}, stdout={result.stdout}, stderr={result.stderr}"
        )

    if not expected_docx.exists():
        raise RuntimeError(
            "LibreOffice 未生成 DOCX 文件: "
            f"stdout={result.stdout}, stderr={result.stderr}"
        )

    if expected_docx != output_file:
        expected_docx.replace(output_file)

    return str(output_file)


def convert_doc_to_docx_via_libreoffice(
    input_path: str | Path,
    output_path: str | Path | None = None,
    *,
    libreoffice_path: str | None = None,
) -> str:
    input_file = Path(input_path).resolve()
    if input_file.suffix.lower() not in {".doc", ".docx"}:
        raise ValueError(f"仅支持 .doc / .docx 输入，当前为: {input_file.name}")
    return convert_to_docx_via_libreoffice(
        input_path=input_file,
        output_path=output_path,
        libreoffice_path=libreoffice_path,
    )
