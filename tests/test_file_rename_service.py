import json
from pathlib import Path
from types import SimpleNamespace

import anyio
import pytest

from app.service import file_rename_service
from app.controller import task as task_controller


def _allow_root(monkeypatch, root: Path) -> None:
    monkeypatch.setattr(
        file_rename_service.settings,
        "WORD_COUNT_ALLOWED_ROOTS_JSON",
        json.dumps([str(root)], ensure_ascii=False),
    )
    monkeypatch.setattr(file_rename_service.settings, "WORD_COUNT_UNC_MOUNT_MAP_JSON", "")
    monkeypatch.setattr(file_rename_service.settings, "WORD_COUNT_ALLOW_LOCAL_PATHS", "False")
    monkeypatch.setattr(file_rename_service.settings, "WORD_COUNT_FOLLOW_SYMLINKS", "False")


def test_discover_reuses_allowed_path_and_skips_generated_copy_directories(monkeypatch, tmp_path):
    root = tmp_path / "共享目录"
    nested = root / "子目录"
    generated = root / f"{file_rename_service.FILE_RENAME_COPY_DIR_PREFIX}旧任务"
    nested.mkdir(parents=True)
    generated.mkdir()
    (root / "b.pdf").write_bytes(b"pdf")
    (root / "a.txt").write_text("甲", encoding="utf-8")
    (nested / "c.docx").write_bytes(b"docx")
    (root / ".hidden.txt").write_text("hidden", encoding="utf-8")
    (root / "~$lock.docx").write_bytes(b"lock")
    (generated / "old.txt").write_text("old", encoding="utf-8")
    _allow_root(monkeypatch, root)

    result = file_rename_service.discover_file_rename_files(
        directory_path=str(root),
        recursive=True,
        include_hidden=False,
    )

    assert [item["relative_path"] for item in result["files"]] == [
        "a.txt",
        "b.pdf",
        "子目录/c.docx",
    ]
    assert result["truncated"] is False


def test_numbering_preview_uses_one_relative_path_sorted_sequence(monkeypatch, tmp_path):
    root = tmp_path / "项目"
    nested = root / "子目录"
    nested.mkdir(parents=True)
    (root / "b.pdf").write_bytes(b"b")
    (root / "a.txt").write_bytes(b"a")
    (nested / "c.docx").write_bytes(b"c")
    _allow_root(monkeypatch, root)

    preview = file_rename_service.build_file_rename_preview(
        directory_path=str(root),
        relative_paths=["子目录/c.docx", "b.pdf", "a.txt"],
        mode="numbering",
    )

    assert preview["process_count"] == 3
    assert preview["number_width"] == 1
    assert [item["target_relative_path"] for item in preview["operations"]] == [
        "1_a.txt",
        "2_b.pdf",
        "子目录/3_c.docx",
    ]

    prepared = file_rename_service.prepare_file_rename_request(
        directory_path=str(root),
        relative_paths=["子目录/c.docx", "b.pdf", "a.txt"],
        mode="numbering",
    )
    assert prepared["params"]["directory_path"] == str(root)
    assert prepared["params"]["relative_paths"] == ["a.txt", "b.pdf", "子目录/c.docx"]


def test_regex_preview_changes_stem_and_preserves_extension(monkeypatch, tmp_path):
    root = tmp_path / "项目"
    root.mkdir()
    source = root / "2026.2.3_xxxx_translated.docx"
    source.write_bytes(b"source")
    (root / "普通文件.docx").write_bytes(b"plain")
    _allow_root(monkeypatch, root)

    preview = file_rename_service.build_file_rename_preview(
        directory_path=str(root),
        relative_paths=[source.name, "普通文件.docx"],
        mode="regex",
        regex_pattern=r"^\d{4}\.\d{1,2}\.\d{1,2}_(.*?)_translated$",
        replacement=r"\1",
    )

    assert preview["process_count"] == 1
    assert preview["skipped_count"] == 1
    assert preview["operations"][0]["target_relative_path"] == "xxxx.docx"
    assert preview["operations"][1]["status"] == "unmatched"


def test_common_cleanup_handles_business_filename_examples(monkeypatch, tmp_path):
    root = tmp_path / "项目"
    root.mkdir()
    filenames = [
        "01_20160907-Homart蓝莓和其他产品研发费用.pdf",
        "004 20240131功能清单备注_translated.docx",
        "20260702-000853_IMG_6363.docx",
        "2026.2.3_xxxx_translated.docx",
        "普通文件.docx",
    ]
    for filename in filenames:
        (root / filename).write_bytes(filename.encode("utf-8"))
    _allow_root(monkeypatch, root)

    preview = file_rename_service.build_file_rename_preview(
        directory_path=str(root),
        relative_paths=filenames,
        mode="cleanup",
    )
    operations = {
        item["source_relative_path"]: item
        for item in preview["operations"]
    }

    assert operations[filenames[0]]["target_relative_path"] == "20160907-Homart蓝莓和其他产品研发费用.pdf"
    assert operations[filenames[1]]["target_relative_path"] == "20240131功能清单备注.docx"
    assert operations[filenames[2]]["target_relative_path"] == "IMG_6363.docx"
    assert operations[filenames[3]]["target_relative_path"] == "xxxx.docx"
    assert operations[filenames[4]]["status"] == "unchanged"
    assert preview["process_count"] == 4
    assert preview["skipped_count"] == 1


def test_common_cleanup_options_can_be_fine_tuned(monkeypatch, tmp_path):
    root = tmp_path / "项目"
    root.mkdir()
    source = root / "123_文件_END.pdf"
    source.write_bytes(b"source")
    _allow_root(monkeypatch, root)

    preview = file_rename_service.build_file_rename_preview(
        directory_path=str(root),
        relative_paths=[source.name],
        mode="cleanup",
        cleanup_leading_number_max_digits=3,
        cleanup_leading_number_space=False,
        cleanup_leading_number_underscore=True,
        cleanup_remove_datetime=False,
        cleanup_translated_suffix="_END",
    )
    assert preview["operations"][0]["target_relative_path"] == "文件.pdf"

    unchanged = file_rename_service.build_file_rename_preview(
        directory_path=str(root),
        relative_paths=[source.name],
        mode="cleanup",
        cleanup_leading_number_space=True,
        cleanup_leading_number_underscore=False,
        cleanup_remove_datetime=False,
        cleanup_remove_translated=False,
    )
    assert unchanged["operations"][0]["status"] == "unchanged"


def test_regex_preview_blocks_duplicate_targets(monkeypatch, tmp_path):
    root = tmp_path / "项目"
    root.mkdir()
    (root / "a.txt").write_bytes(b"a")
    (root / "b.txt").write_bytes(b"b")
    _allow_root(monkeypatch, root)

    preview = file_rename_service.build_file_rename_preview(
        directory_path=str(root),
        relative_paths=["a.txt", "b.txt"],
        mode="regex",
        regex_pattern=r"^.*$",
        replacement="same",
    )

    assert preview["process_count"] == 0
    assert preview["conflict_count"] == 2
    with pytest.raises(ValueError, match="重名冲突"):
        file_rename_service.prepare_file_rename_request(
            directory_path=str(root),
            relative_paths=["a.txt", "b.txt"],
            mode="regex",
            regex_pattern=r"^.*$",
            replacement="same",
        )


def test_execute_creates_renamed_copy_and_keeps_sources(monkeypatch, tmp_path):
    root = tmp_path / "项目"
    nested = root / "子目录"
    nested.mkdir(parents=True)
    first = root / "a.txt"
    second = nested / "b.txt"
    first.write_text("alpha", encoding="utf-8")
    second.write_text("beta", encoding="utf-8")
    _allow_root(monkeypatch, root)
    progress_updates = []

    async def run_task():
        async def update(progress, message):
            progress_updates.append((progress, message))

        return await file_rename_service.execute_file_rename_copy_task(
            task_id="task-1",
            display_no="TASK-1",
            directory_path=str(root),
            relative_paths=["子目录/b.txt", "a.txt"],
            mode="numbering",
            progress_callback=update,
        )

    result = anyio.run(run_task)
    output_directory = Path(result["output_directory"])

    assert first.read_text(encoding="utf-8") == "alpha"
    assert second.read_text(encoding="utf-8") == "beta"
    assert (output_directory / "1_a.txt").read_text(encoding="utf-8") == "alpha"
    assert (output_directory / "子目录" / "2_b.txt").read_text(encoding="utf-8") == "beta"
    assert result["copied_count"] == 2
    assert result["failed_count"] == 0
    assert progress_updates[-1][0] == 100


def test_file_rename_preview_and_submit_endpoints(monkeypatch, tmp_path):
    root = tmp_path / "项目"
    root.mkdir()
    (root / "a.txt").write_bytes(b"a")
    _allow_root(monkeypatch, root)

    body = task_controller.FileRenameRequestBody(
        directory_path=str(root),
        relative_paths=["a.txt"],
        mode="numbering",
    )
    preview = anyio.run(task_controller.preview_file_rename, body)
    assert preview["operations"][0]["target_relative_path"] == "1_a.txt"

    captured = {}

    async def fake_submit_file_rename_task(**kwargs):
        captured.update(kwargs)
        return SimpleNamespace(task_id="rename-task", deduped=False)

    monkeypatch.setattr(
        task_controller.task_queue_service,
        "submit_file_rename_task",
        fake_submit_file_rename_task,
    )
    result = anyio.run(task_controller.submit_file_rename, body)

    assert result["task_id"] == "rename-task"
    assert captured["directory_path"] == str(root)
    assert captured["relative_paths"] == ["a.txt"]
    assert captured["mode"] == "numbering"


def test_file_rename_page_exposes_full_scan_csv_export():
    project_root = Path(__file__).resolve().parents[1]
    html = (project_root / "static" / "file_rename.html").read_text(encoding="utf-8")
    javascript = (project_root / "static" / "file_rename.js").read_text(encoding="utf-8")

    assert 'id="exportListBtn"' in html
    assert "导出文件名单" in html
    assert "function exportFileList()" in javascript
    assert "['序号', '文件名', '相对路径', '扩展名', '文件大小（字节）', '修改时间']" in javascript
    assert "!scanCompleted || !discoveredFiles.length" in javascript
    assert "\\uFEFF" in javascript


def test_file_rename_page_renders_explorer_style_folder_tree():
    project_root = Path(__file__).resolve().parents[1]
    html = (project_root / "static" / "file_rename.html").read_text(encoding="utf-8")
    javascript = (project_root / "static" / "file_rename.js").read_text(encoding="utf-8")

    assert 'id="expandAllBtn"' in html
    assert 'id="collapseAllBtn"' in html
    assert "function buildFileTree(files)" in javascript
    assert "data-folder-select" in javascript
    assert "input.indeterminate" in javascript
    assert "function setAllFoldersExpanded(expanded)" in javascript


def test_file_rename_cleanup_controls_keep_user_friendly_defaults():
    project_root = Path(__file__).resolve().parents[1]
    html = (project_root / "static" / "file_rename.html").read_text(encoding="utf-8")
    javascript = (project_root / "static" / "file_rename.js").read_text(encoding="utf-8")

    assert html.index('data-mode="numbering"') < html.index('data-mode="cleanup"')
    assert '<button class="mode-button active" data-mode="cleanup">' in html
    assert 'id="cleanupMaxDigitsInput"' in html
    assert 'id="cleanupSeparatorUnderscoreInput" type="checkbox" checked' in html
    assert 'id="cleanupSuffixInput"' in html
    assert "function updateCleanupControlAvailability()" in javascript
