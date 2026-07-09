import io
import json
import sys
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZIP_DEFLATED, ZipFile

import anyio
import pytest
from docx import Document
from openpyxl import Workbook, load_workbook
from pptx import Presentation
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.controller import task as task_controller
from app.db.database import Base
from app.model.entity import Task
from app.service import task_queue_service as queue_module
from app.service import word_count_service


def _make_png(path: Path) -> None:
    from PIL import Image

    image = Image.new("RGB", (24, 24), color=(48, 128, 220))
    image.save(path)


def _set_docx_app_props(path: Path, updates: dict[str, str]) -> None:
    temp_path = path.with_suffix(".tmp.docx")
    with ZipFile(path, "r") as source, ZipFile(temp_path, "w", ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "docProps/app.xml":
                from lxml import etree

                root = etree.fromstring(data)
                for child in root:
                    tag = etree.QName(child).localname
                    if tag in updates:
                        child.text = str(updates[tag])
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            target.writestr(item, data)
    path.unlink()
    temp_path.rename(path)


def _allow_root(monkeypatch, root: Path):
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOWED_ROOTS_JSON", json.dumps([str(root)]))
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_UNC_MOUNT_MAP_JSON", "")
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "False")
    monkeypatch.setattr(queue_module.settings, "WORD_COUNT_ALLOWED_ROOTS_JSON", json.dumps([str(root)]))
    monkeypatch.setattr(queue_module.settings, "WORD_COUNT_UNC_MOUNT_MAP_JSON", "")
    monkeypatch.setattr(queue_module.settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "False")


def test_word_like_counter_counts_mixed_languages():
    text = "你好 world 123，テスト 한국어"

    assert word_count_service.count_words_word_like(text) == 11


def test_word_like_counter_matches_word_stats_sample():
    sample = Path("data/word/2. Blog Post 42 - ENG - Red Notice removal from INTERPOL, and when NOT to apply - word_translated.docx")
    if not sample.exists():
        pytest.skip("sample DOCX is not available")

    items = word_count_service._extract_docx_text_items(sample)
    main_text = "\n".join(item.text for item in items if not item.is_extra)
    metrics = word_count_service._count_text(main_text)

    assert metrics.word_count == 693
    assert metrics.non_space_chars == 724


def test_word_like_counter_handles_edge_punctuation():
    assert word_count_service.count_words_word_like("“中文”") == 4
    assert word_count_service.count_words_word_like("中文…") == 3


def test_docx_main_and_extra_counts_are_separated(tmp_path, monkeypatch):
    _allow_root(monkeypatch, tmp_path)
    monkeypatch.setattr(word_count_service.settings, "OUTPUT_DIR", str(tmp_path / "outputs"))

    image_path = tmp_path / "doc-image.png"
    _make_png(image_path)
    doc = Document()
    doc.add_paragraph("你好 world")
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "金额 123"
    section = doc.sections[0]
    section.header.paragraphs[0].text = "页眉 text"
    doc.add_picture(str(image_path))
    doc.save(tmp_path / "sample.docx")
    _set_docx_app_props(tmp_path / "sample.docx", {"Pages": "2", "Lines": "9"})

    result = word_count_service.run_word_count_task_sync(
        task_id="task-1",
        display_no="000001",
        directory_path=str(tmp_path),
        recursive=True,
        include_hidden=False,
        extensions=[".docx"],
    )

    summary = result["summary"]
    assert summary["counted_files"] == 1
    assert summary["total_main_word_count"] == 6
    assert summary["total_extra_word_count"] == 3
    assert summary["total_page_count"] == 2
    assert summary["total_line_count"] == 9
    assert summary["total_image_count"] == 1
    file_result = result["files"][0]
    assert file_result["word_count"] == 6
    assert file_result["char_count_no_spaces"] > 0
    assert file_result["char_count_with_spaces"] >= file_result["char_count_no_spaces"]
    assert file_result["paragraph_count"] >= 3
    assert file_result["image_count"] == 1
    assert file_result["file_type"] == "Word"
    assert "DOCX" in file_result["stat_method"]
    assert file_result["counted_at"]
    assert Path(tmp_path / result["report_excel"]).exists()
    assert Path(tmp_path / result["report_json"]).exists()

    report = load_workbook(tmp_path / result["report_excel"], read_only=True)
    try:
        headers = [cell.value for cell in next(report["文件明细"].iter_rows(max_row=1))]
        assert "页数" in headers
        assert "字符数(不计空格)" in headers
        assert "图片数量" in headers
        assert "统计方法" in headers
    finally:
        report.close()


def test_office_pdf_txt_and_blank_pdf_statuses(tmp_path, monkeypatch):
    _allow_root(monkeypatch, tmp_path)
    monkeypatch.setattr(word_count_service.settings, "OUTPUT_DIR", str(tmp_path / "outputs"))

    image_path = tmp_path / "embedded.png"
    _make_png(image_path)
    txt_path = tmp_path / "plain.txt"
    txt_path.write_text("中文 text 42\nsecond line", encoding="utf-8")

    workbook = Workbook()
    workbook.active["A1"] = "表格 abc"
    workbook.active["A2"] = "第二行"
    from openpyxl.drawing.image import Image as XlsxImage

    workbook.active.add_image(XlsxImage(str(image_path)), "B2")
    workbook.create_sheet("空表")
    workbook.save(tmp_path / "sheet.xlsx")

    presentation = Presentation()
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "标题 one"
    tx_box = slide.shapes.add_textbox(0, 0, 914400, 914400)
    tx_box.text_frame.text = "文本框 two"
    table = slide.shapes.add_table(1, 1, 0, 914400, 914400, 914400).table
    table.cell(0, 0).text = "表格 three"
    slide.shapes.add_picture(str(image_path), 0, 1828800)
    slide.notes_slide.notes_text_frame.text = "备注 extra"
    presentation.save(tmp_path / "deck.pptx")

    import fitz

    pdf = fitz.open()
    page = pdf.new_page()
    page.insert_text((72, 72), "PDF text")
    pdf.save(tmp_path / "text.pdf")
    pdf.close()

    blank_pdf = fitz.open()
    blank_pdf.new_page()
    blank_pdf.save(tmp_path / "blank.pdf")
    blank_pdf.close()

    result = word_count_service.run_word_count_task_sync(
        task_id="task-2",
        display_no="000002",
        directory_path=str(tmp_path),
        recursive=True,
        include_hidden=False,
        extensions=[".txt", ".xlsx", ".pptx", ".pdf"],
    )

    statuses = {item["filename"]: item["status"] for item in result["files"]}
    assert statuses["plain.txt"] == "counted"
    assert statuses["sheet.xlsx"] == "counted"
    assert statuses["deck.pptx"] == "counted"
    assert statuses["text.pdf"] == "counted"
    assert statuses["blank.pdf"] == "needs_ocr"
    assert result["summary"]["needs_ocr_files"] == 1
    files = {item["filename"]: item for item in result["files"]}
    assert files["plain.txt"]["page_count"] == 1
    assert files["plain.txt"]["line_count"] == 2
    assert files["sheet.xlsx"]["page_count"] == 2
    assert files["sheet.xlsx"]["line_count"] == 2
    assert files["sheet.xlsx"]["image_count"] == 1
    assert files["sheet.xlsx"]["file_type"] == "Excel"
    assert files["deck.pptx"]["page_count"] == 1
    assert files["deck.pptx"]["image_count"] == 1
    assert files["deck.pptx"]["extra_word_count"] > 0
    assert files["deck.pptx"]["file_type"] == "PPT"
    assert files["text.pdf"]["page_count"] == 1
    assert files["blank.pdf"]["page_count"] == 1
    assert result["summary"]["total_image_count"] >= 2


def test_directory_must_be_inside_allowed_roots(tmp_path, monkeypatch):
    allowed = tmp_path / "allowed"
    outside = tmp_path / "outside"
    allowed.mkdir()
    outside.mkdir()
    _allow_root(monkeypatch, allowed)

    with pytest.raises(PermissionError):
        word_count_service.prepare_word_count_request(directory_path=str(outside))


def test_local_paths_can_be_allowed_for_testing(tmp_path, monkeypatch):
    allowed = tmp_path / "allowed"
    outside = tmp_path / "outside"
    allowed.mkdir()
    outside.mkdir()
    _allow_root(monkeypatch, allowed)
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "True")

    result = word_count_service.prepare_word_count_request(directory_path=str(outside))

    assert Path(result["params"]["directory_path"]) == outside.resolve(strict=False)


def test_unc_server_root_allows_shares_on_same_server():
    root = Path("\\\\win-server\\")

    assert word_count_service._is_relative_to_path(Path("\\\\win-server\\服务器资料3\\项目"), root)
    assert not word_count_service._is_relative_to_path(Path("\\\\other-server\\服务器资料3\\项目"), root)


def test_unc_path_maps_to_container_mount_before_resolving(tmp_path, monkeypatch):
    mount_root = tmp_path / "mnt" / "win-server" / "服务器资料7"
    target_dir = mount_root / "客户" / "其他客户翻译任务" / "2026年" / "7月" / "美国玲翻译" / "0707" / "1. 原文"
    target_dir.mkdir(parents=True)
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOWED_ROOTS_JSON", json.dumps(["\\\\win-server\\服务器资料7\\"]))
    monkeypatch.setattr(
        word_count_service.settings,
        "WORD_COUNT_UNC_MOUNT_MAP_JSON",
        json.dumps({"\\\\win-server\\服务器资料7\\": str(mount_root)}, ensure_ascii=False),
    )
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "False")

    result = word_count_service.prepare_word_count_request(
        directory_path="\\\\win-server\\服务器资料7\\客户\\其他客户翻译任务\\2026年\\7月\\美国玲翻译\\0707\\1. 原文"
    )

    assert Path(result["params"]["directory_path"]) == target_dir.resolve(strict=False)
    assert Path(result["input_files"]["allowed_root"]) == mount_root.resolve(strict=False)


def test_unc_path_auto_maps_common_container_mount(tmp_path, monkeypatch):
    auto_root = tmp_path / "mnt"
    mount_root = auto_root / "win-server" / "服务器资料7"
    target_dir = mount_root / "客户" / "项目"
    target_dir.mkdir(parents=True)
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOWED_ROOTS_JSON", json.dumps(["\\\\win-server\\服务器资料7\\"]))
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_UNC_MOUNT_MAP_JSON", "")
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_UNC_AUTO_MOUNT_ROOTS_JSON", json.dumps([str(auto_root)], ensure_ascii=False))
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "False")

    result = word_count_service.prepare_word_count_request(directory_path="\\\\win-server\\服务器资料7\\客户\\项目")

    assert Path(result["params"]["directory_path"]) == target_dir.resolve(strict=False)
    assert Path(result["input_files"]["allowed_root"]) == mount_root.resolve(strict=False)


def test_unc_mapped_missing_path_reports_visible_mount_contents(tmp_path, monkeypatch):
    mount_root = tmp_path / "mnt" / "win-server" / "服务器资料7"
    (mount_root / "实际可见目录").mkdir(parents=True)
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOWED_ROOTS_JSON", json.dumps(["\\\\win-server\\服务器资料7\\"]))
    monkeypatch.setattr(
        word_count_service.settings,
        "WORD_COUNT_UNC_MOUNT_MAP_JSON",
        json.dumps({"\\\\win-server\\服务器资料7\\": str(mount_root)}, ensure_ascii=False),
    )
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "False")

    with pytest.raises(FileNotFoundError) as exc_info:
        word_count_service.prepare_word_count_request(directory_path="\\\\win-server\\服务器资料7\\客户\\项目")

    message = str(exc_info.value)
    assert str(mount_root.resolve(strict=False)) in message
    assert "实际可见目录" in message
    assert "Docker 当前挂载到了本地空目录" in message


def test_unc_path_error_lists_auto_mount_candidates(monkeypatch):
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_UNC_AUTO_MOUNT_ROOTS_JSON", json.dumps(["/mnt"], ensure_ascii=False))

    message = word_count_service._format_unc_auto_candidates("\\\\win-server\\服务器资料7\\客户\\项目")

    assert "/mnt/win-server/服务器资料7/客户/项目" in message.replace("\\", "/")
    assert "/mnt/服务器资料7/客户/项目" in message.replace("\\", "/")


def test_unc_mapping_still_requires_allowed_root(tmp_path, monkeypatch):
    mount_root = tmp_path / "mnt" / "win-server" / "服务器资料7"
    target_dir = mount_root / "客户"
    target_dir.mkdir(parents=True)
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOWED_ROOTS_JSON", json.dumps(["\\\\other-server\\share\\"]))
    monkeypatch.setattr(
        word_count_service.settings,
        "WORD_COUNT_UNC_MOUNT_MAP_JSON",
        json.dumps({"\\\\win-server\\服务器资料7\\": str(mount_root)}, ensure_ascii=False),
    )
    monkeypatch.setattr(word_count_service.settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "False")

    with pytest.raises(PermissionError):
        word_count_service.prepare_word_count_request(directory_path="\\\\win-server\\服务器资料7\\客户")


def test_word_count_config_and_submit_endpoint(monkeypatch, tmp_path):
    _allow_root(monkeypatch, tmp_path)

    config = anyio.run(task_controller.get_word_count_page_config)
    assert config["allowed_roots"][0]["path"] == str(tmp_path)
    assert ".docx" in config["countable_extensions"]

    captured = {}

    async def fake_submit_word_count_task(**kwargs):
        captured.update(kwargs)
        return SimpleNamespace(task_id="task-1", deduped=False)

    monkeypatch.setattr(task_controller.task_queue_service, "submit_word_count_task", fake_submit_word_count_task)

    body = task_controller.WordCountSubmitBody(directory_path=str(tmp_path), recursive=False, include_hidden=True, extensions=["docx"])
    result = anyio.run(task_controller.submit_word_count, body)

    assert result["status"] == "ACCEPTED"
    assert captured["directory_path"] == str(tmp_path)
    assert captured["recursive"] is False
    assert captured["include_hidden"] is True
    assert captured["extensions"] == ["docx"]


def test_word_count_task_submission_uses_queue_dedupe(tmp_path, monkeypatch):
    _allow_root(monkeypatch, tmp_path)
    engine = create_engine(
        f"sqlite:///{tmp_path / 'tasks.db'}",
        connect_args={"check_same_thread": False},
    )
    Base.metadata.create_all(bind=engine)
    with engine.begin() as connection:
        connection.execute(
            text(
                "CREATE UNIQUE INDEX ux_task_active_request_fingerprint "
                "ON task (request_fingerprint) "
                "WHERE request_fingerprint IS NOT NULL AND status IN ('queued', 'running')"
            )
        )

    testing_session = sessionmaker(autocommit=False, autoflush=False, bind=engine)
    monkeypatch.setattr(queue_module, "SessionLocal", testing_session)

    service = queue_module.TaskQueueService()

    async def scenario():
        first = await service.submit_word_count_task(directory_path=str(tmp_path))
        second = await service.submit_word_count_task(directory_path=str(tmp_path))
        return first, second

    first, second = anyio.run(scenario)

    assert first.task_id == second.task_id
    assert first.deduped is False
    assert second.deduped is True
    with testing_session() as db:
        task = db.query(Task).one()
        assert task.task_type == "word_count"
        assert json.loads(task.input_files_json)["directory_path"] == str(tmp_path.resolve())
