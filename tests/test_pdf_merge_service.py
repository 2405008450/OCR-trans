import asyncio
from pathlib import Path

import pytest
from pypdf import PdfReader, PdfWriter

from app.core.config import settings
from app.service.pdf_merge_service import (
    discover_pdf_files,
    execute_pdf_merge_task,
    normalize_output_filename,
    prepare_pdf_merge_request,
)


def _write_pdf(path: Path, page_widths: list[float]) -> None:
    writer = PdfWriter()
    for width in page_widths:
        writer.add_blank_page(width=width, height=400)
    with path.open("wb") as output:
        writer.write(output)
    writer.close()


@pytest.fixture
def local_shared_root(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    shared_root = tmp_path / "共享目录"
    shared_root.mkdir()
    monkeypatch.setattr(settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "True")
    monkeypatch.setattr(settings, "OUTPUT_DIR", str(tmp_path / "outputs"))
    monkeypatch.setattr(settings, "UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setattr(settings, "PDF_MERGE_MAX_FILES", 20)
    monkeypatch.setattr(settings, "PDF_MERGE_MAX_FILE_MB", 10)
    monkeypatch.setattr(settings, "PDF_MERGE_MAX_TOTAL_MB", 20)
    return shared_root


def test_discover_pdf_files_uses_natural_order(local_shared_root: Path) -> None:
    _write_pdf(local_shared_root / "10.pdf", [210])
    _write_pdf(local_shared_root / "2.pdf", [220])
    _write_pdf(local_shared_root / ".hidden.pdf", [230])
    subdirectory = local_shared_root / "子目录"
    subdirectory.mkdir()
    _write_pdf(subdirectory / "3.pdf", [240])

    result = discover_pdf_files(directory_path=str(local_shared_root), recursive=True)

    assert [item["relative_path"] for item in result["files"]] == [
        "2.pdf",
        "10.pdf",
        "子目录/3.pdf",
    ]
    assert [item["page_count"] for item in result["files"]] == [1, 1, 1]
    assert result["truncated"] is False


def test_discover_pdf_files_keeps_unreadable_pdf_with_unknown_page_count(local_shared_root: Path) -> None:
    (local_shared_root / "broken.pdf").write_bytes(b"not a valid PDF")

    result = discover_pdf_files(directory_path=str(local_shared_root), recursive=False)

    assert result["files"][0]["relative_path"] == "broken.pdf"
    assert result["files"][0]["page_count"] is None


def test_prepare_pdf_merge_rejects_invalid_selection(local_shared_root: Path) -> None:
    _write_pdf(local_shared_root / "a.pdf", [210])
    _write_pdf(local_shared_root / "b.pdf", [220])

    with pytest.raises(ValueError, match="至少选择 2 个"):
        prepare_pdf_merge_request(
            directory_path=str(local_shared_root),
            relative_paths=["a.pdf"],
            output_filename="结果",
        )

    with pytest.raises(ValueError, match="相对路径无效"):
        prepare_pdf_merge_request(
            directory_path=str(local_shared_root),
            relative_paths=["a.pdf", "../b.pdf"],
            output_filename="结果",
        )


def test_execute_pdf_merge_preserves_selected_order_and_page_count(local_shared_root: Path) -> None:
    _write_pdf(local_shared_root / "first.pdf", [210])
    _write_pdf(local_shared_root / "second.pdf", [320, 330])
    progress_updates: list[tuple[int, str]] = []

    async def progress(progress_value: int, message: str) -> None:
        progress_updates.append((progress_value, message))

    result = asyncio.run(
        execute_pdf_merge_task(
            task_id="test-task",
            display_no="PDF-001",
            directory_path=str(local_shared_root),
            relative_paths=["second.pdf", "first.pdf"],
            output_filename="客户合并件",
            progress_callback=progress,
        )
    )

    output_path = Path(settings.OUTPUT_DIR) / "pdf_merge" / "PDF-001" / "客户合并件.pdf"
    reader = PdfReader(output_path)
    widths = [round(float(page.mediabox.width)) for page in reader.pages]

    assert result["output_pdf"] == "outputs/pdf_merge/PDF-001/客户合并件.pdf"
    assert result["input_file_count"] == 2
    assert result["total_pages"] == 3
    assert widths == [320, 330, 210]
    assert [item["relative_path"] for item in result["files"]] == ["second.pdf", "first.pdf"]
    assert progress_updates[-1][0] == 96
    assert not (Path(settings.UPLOAD_DIR) / "pdf_merge_staging" / "PDF-001").exists()


@pytest.mark.parametrize(
    ("raw_name", "expected"),
    [
        ("项目合并件", "项目合并件.pdf"),
        ("../危险:名称.PDF", "危险_名称.pdf"),
        ("", "合并结果.pdf"),
    ],
)
def test_normalize_output_filename(raw_name: str, expected: str) -> None:
    assert normalize_output_filename(raw_name) == expected
