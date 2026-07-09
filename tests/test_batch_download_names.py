import json
import sys
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.controller.task import (
    _build_archive_name,
    _build_batch_archive_download_name,
    _is_legacy_default_archive_name,
)


def _task(filename="contract.pdf", status="done"):
    return SimpleNamespace(
        display_no="20260703-000123",
        task_id="task-abcdef123456",
        status=status,
        filename=filename,
        input_files_json=json.dumps({"original_filename": filename}, ensure_ascii=False),
    )


def test_archive_entry_name_keeps_output_display_name_without_task_prefix():
    used_names = set()

    assert _build_archive_name(_task(), "contract.docx", used_names) == "contract.docx"


def test_archive_entry_name_strips_existing_display_no_prefix():
    used_names = set()

    assert _build_archive_name(_task(), "20260703-000123_contract.docx", used_names) == "contract.docx"


def test_archive_entry_name_only_suffixes_duplicates():
    used_names = set()

    assert _build_archive_name(_task(), "contract.docx", used_names) == "contract.docx"
    assert _build_archive_name(_task(), "contract.docx", used_names) == "contract_2.docx"


def test_batch_archive_download_name_uses_original_filename():
    archive_name = _build_batch_archive_download_name([
        _task("contract.pdf"),
        _task("invoice.pdf"),
    ])

    assert archive_name != "batch_outputs.zip"
    assert archive_name.startswith("contract_")
    assert archive_name.endswith(".zip")


def test_legacy_batch_outputs_name_is_treated_as_default():
    assert _is_legacy_default_archive_name(" batch_outputs.zip ")
