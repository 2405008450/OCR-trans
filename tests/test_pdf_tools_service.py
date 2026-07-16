import asyncio
from pathlib import Path
from zipfile import ZipFile

import fitz
import pytest
from PIL import Image
from pypdf import PdfReader, PdfWriter

from app.core.config import settings
from app.service.pdf_tools_service import (
    execute_pdf_tools_task,
    get_pdf_tools_config,
    parse_page_spec,
    prepare_pdf_tools_request,
)


def _write_pdf(path: Path, page_widths: list[float]) -> None:
    writer = PdfWriter()
    for width in page_widths:
        writer.add_blank_page(width=width, height=400)
    with path.open("wb") as output:
        writer.write(output)
    writer.close()


def _write_image_pdf(path: Path, image_path: Path) -> None:
    image = Image.effect_noise((1800, 2400), 80).convert("RGB")
    image.save(image_path, format="PNG")
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    page.insert_image(page.rect, filename=str(image_path))
    document.save(path)
    document.close()


@pytest.fixture
def local_pdf_root(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    root = tmp_path / "共享 PDF"
    root.mkdir()
    monkeypatch.setattr(settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "True")
    monkeypatch.setattr(settings, "OUTPUT_DIR", str(tmp_path / "outputs"))
    monkeypatch.setattr(settings, "UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setattr(settings, "PDF_MERGE_MAX_FILES", 20)
    monkeypatch.setattr(settings, "PDF_MERGE_MAX_FILE_MB", 50)
    return root


def test_page_spec_parser_supports_ranges_and_all() -> None:
    assert parse_page_spec("1-3,5,3", 6) == [0, 1, 2, 4]
    assert parse_page_spec("all", 3) == [0, 1, 2]
    with pytest.raises(ValueError, match="超出范围"):
        parse_page_spec("1-9", 3)


def test_config_exposes_four_compression_levels() -> None:
    config = get_pdf_tools_config()
    modes = config["compression_modes"]
    assert list(modes) == ["lossless", "high", "balanced", "strong"]
    assert [item["value"] for item in config["page_selection_modes"]] == ["custom", "odd", "even"]


def test_compression_defaults_to_readability_first(local_pdf_root: Path) -> None:
    _write_pdf(local_pdf_root / "source.pdf", [200])
    prepared = prepare_pdf_tools_request(
        directory_path=str(local_pdf_root),
        relative_path="source.pdf",
        operation="compress",
        options={},
    )
    assert prepared["params"]["options"]["compression_mode"] == "high"


def test_prepare_rejects_path_traversal(local_pdf_root: Path) -> None:
    _write_pdf(local_pdf_root / "source.pdf", [200, 210])
    with pytest.raises(ValueError, match="相对路径无效"):
        prepare_pdf_tools_request(
            directory_path=str(local_pdf_root),
            relative_path="../source.pdf",
            operation="split",
            options={"pages_per_file": 1},
        )


def test_split_pdf_creates_ordered_parts_and_zip(local_pdf_root: Path) -> None:
    _write_pdf(local_pdf_root / "source.pdf", [200, 210, 220, 230, 240])
    result = asyncio.run(
        execute_pdf_tools_task(
            task_id="split-task",
            display_no="PDF-TOOLS-001",
            directory_path=str(local_pdf_root),
            relative_path="source.pdf",
            operation="split",
            options={"split_mode": "every", "pages_per_file": 2, "output_prefix": "客户文件"},
        )
    )

    assert result["source_page_count"] == 5
    assert result["output_file_count"] == 3
    assert [item["page_count"] for item in result["output_files"]] == [2, 2, 1]
    output_dir = Path(settings.OUTPUT_DIR) / "pdf_tools" / "PDF-TOOLS-001"
    archive_path = output_dir / result["archive_filename"]
    with ZipFile(archive_path) as archive:
        assert archive.namelist() == [item["filename"] for item in result["output_files"]]
    assert not (Path(settings.UPLOAD_DIR) / "pdf_tools_staging" / "PDF-TOOLS-001").exists()


@pytest.mark.parametrize(
    ("operation", "options", "expected_pages"),
    [
        ("extract", {"page_spec": "2-3", "output_filename": "提取.pdf"}, 2),
        ("delete", {"page_spec": "2", "output_filename": "删页.pdf"}, 3),
        ("rotate", {"page_spec": "1,3", "angle": 90, "output_filename": "旋转.pdf"}, 4),
    ],
)
def test_page_operations(local_pdf_root: Path, operation: str, options: dict, expected_pages: int) -> None:
    _write_pdf(local_pdf_root / "pages.pdf", [200, 210, 220, 230])
    result = asyncio.run(
        execute_pdf_tools_task(
            task_id=f"{operation}-task",
            display_no=f"PDF-{operation}",
            directory_path=str(local_pdf_root),
            relative_path="pages.pdf",
            operation=operation,
            options=options,
        )
    )
    output_path = Path(settings.OUTPUT_DIR) / "pdf_tools" / f"PDF-{operation}" / options["output_filename"]
    reader = PdfReader(output_path)
    assert len(reader.pages) == expected_pages
    assert result["total_pages"] == expected_pages
    if operation == "rotate":
        assert [reader.pages[index].rotation for index in range(4)] == [90, 0, 90, 0]


@pytest.mark.parametrize(
    ("operation", "page_mode", "expected_widths", "selected_pages"),
    [
        ("extract", "odd", [200, 220, 240], [1, 3, 5]),
        ("extract", "even", [210, 230], [2, 4]),
        ("delete", "odd", [210, 230], [1, 3, 5]),
        ("delete", "even", [200, 220, 240], [2, 4]),
    ],
)
def test_page_operations_support_odd_and_even_modes(
    local_pdf_root: Path,
    operation: str,
    page_mode: str,
    expected_widths: list[float],
    selected_pages: list[int],
) -> None:
    _write_pdf(local_pdf_root / "odd-even.pdf", [200, 210, 220, 230, 240])
    output_filename = f"{operation}-{page_mode}.pdf"
    result = asyncio.run(
        execute_pdf_tools_task(
            task_id=f"{operation}-{page_mode}-task",
            display_no=f"PDF-{operation}-{page_mode}",
            directory_path=str(local_pdf_root),
            relative_path="odd-even.pdf",
            operation=operation,
            options={"page_mode": page_mode, "output_filename": output_filename},
        )
    )

    output_path = Path(settings.OUTPUT_DIR) / "pdf_tools" / f"PDF-{operation}-{page_mode}" / output_filename
    reader = PdfReader(output_path)
    actual_widths = [float(page.mediabox.width) for page in reader.pages]
    assert actual_widths == expected_widths
    assert result["selected_pages"] == selected_pages
    assert result["page_mode"] == page_mode


def test_extract_even_rejects_document_without_even_pages(local_pdf_root: Path) -> None:
    _write_pdf(local_pdf_root / "one-page.pdf", [200])
    with pytest.raises(ValueError, match="没有可处理的偶数页"):
        asyncio.run(
            execute_pdf_tools_task(
                task_id="extract-even-empty-task",
                display_no="PDF-EXTRACT-EVEN-EMPTY",
                directory_path=str(local_pdf_root),
                relative_path="one-page.pdf",
                operation="extract",
                options={"page_mode": "even", "output_filename": "偶数页.pdf"},
            )
        )


@pytest.mark.parametrize("compression_mode", ["lossless", "high", "balanced", "strong"])
def test_compression_modes_never_return_larger_file(
    local_pdf_root: Path,
    tmp_path: Path,
    compression_mode: str,
) -> None:
    image_pdf = local_pdf_root / f"scan-{compression_mode}.pdf"
    _write_image_pdf(image_pdf, tmp_path / f"noise-{compression_mode}.png")
    result = asyncio.run(
        execute_pdf_tools_task(
            task_id=f"compress-{compression_mode}-task",
            display_no=f"PDF-COMPRESS-{compression_mode}",
            directory_path=str(local_pdf_root),
            relative_path=image_pdf.name,
            operation="compress",
            options={"compression_mode": compression_mode, "output_filename": "压缩.pdf"},
        )
    )
    output_path = Path(settings.OUTPUT_DIR) / "pdf_tools" / f"PDF-COMPRESS-{compression_mode}" / "压缩.pdf"
    assert len(PdfReader(output_path).pages) == 1
    assert result["output_size"] <= result["input_size"]
    assert result["compression_mode"] == compression_mode
