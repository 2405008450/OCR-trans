import inspect
import importlib
import sys
import types
from pathlib import Path

import pytest
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.service import number_check_service


def test_number_check_service_uses_latest_specialist_root():
    expected_root = Path(__file__).resolve().parents[1] / "专检" / "数检_程序-AI"

    assert number_check_service.NUMBER_CHECK_LATEST_ROOT == expected_root
    assert number_check_service.NUMBER_CHECK_MAIN_FILE == expected_root / "main.py"
    assert number_check_service.NUMBER_CHECK_MAIN_FILE.is_file()


def test_number_check_root_supports_configured_relative_path(tmp_path, monkeypatch):
    configured_root = tmp_path / "deployed-number-check"
    configured_root.mkdir()
    (configured_root / "main.py").write_text("def run(): pass\n", encoding="utf-8")
    relative_path = configured_root.relative_to(tmp_path)

    monkeypatch.setattr(number_check_service, "REPO_ROOT", tmp_path)
    monkeypatch.setattr(number_check_service.settings, "NUMBER_CHECK_ROOT", str(relative_path))

    assert number_check_service._resolve_number_check_root() == configured_root.resolve()


def test_number_check_root_falls_back_to_legacy_deployment_name(tmp_path, monkeypatch):
    fallback_root = tmp_path / "专检" / "数检_程序-AIV2"
    fallback_root.mkdir(parents=True)
    (fallback_root / "main.py").write_text("def run(): pass\n", encoding="utf-8")

    monkeypatch.setattr(number_check_service, "REPO_ROOT", tmp_path)
    monkeypatch.setattr(number_check_service.settings, "NUMBER_CHECK_ROOT", "")
    monkeypatch.chdir(tmp_path)

    assert number_check_service._resolve_number_check_root() == fallback_root.resolve()


def test_clear_specialist_module_cache_removes_modules_from_specialist_tree():
    module_name = "number_check_v2_cache_probe"
    module = types.ModuleType(module_name)
    module.__file__ = str(number_check_service.NUMBER_CHECK_LATEST_ROOT / "cache_probe.py")
    sys.modules[module_name] = module

    number_check_service._clear_specialist_module_cache()

    assert module_name not in sys.modules


def test_latest_main_keeps_system_integration_contract(monkeypatch):
    monkeypatch.setenv("API_KEY", "test-key")
    monkeypatch.setenv("OPENAI_API_KEY", "test-key")

    module = number_check_service._load_latest_main_module()
    parameters = inspect.signature(module.run).parameters

    assert Path(module.__file__).resolve().parent == number_check_service.NUMBER_CHECK_LATEST_ROOT.resolve()
    assert {
        "alignment_path",
        "output_dir",
        "src_docx_path",
        "tgt_docx_path",
        "src_hf_path",
        "docx_path",
        "revised_docx_path",
        "revision_author",
        "use_total_normalizer",
        "force_mode_b",
        "ai_check_all",
        "use_legacy_mode",
        "bilingual_mode",
    }.issubset(parameters)


def test_latest_version_rejects_numeric_suffix_false_match(monkeypatch):
    monkeypatch.setenv("API_KEY", "test-key")
    monkeypatch.setenv("OPENAI_API_KEY", "test-key")
    number_check_service._load_latest_main_module()
    replace_revision = importlib.import_module("replace_revision")

    assert replace_revision._find_safe_exact_span("比例为0.00025%", "5%") is None
    match = replace_revision._find_safe_exact_span("增长5%，达到目标", "5%")
    assert match is not None
    assert match.group() == "5%"


def test_docx_revised_output_is_initialized_before_v2_run(tmp_path, monkeypatch):
    target_path = tmp_path / "translated.docx"
    target_content = b"docx-package-placeholder"
    target_path.write_bytes(target_content)
    output_dir = tmp_path / "outputs"
    output_dir.mkdir()
    captured = {}

    def fake_run(**kwargs):
        revised_path = Path(kwargs["revised_docx_path"])
        captured["revised_path"] = revised_path
        assert revised_path.is_file()
        assert revised_path.read_bytes() == target_content
        return [], [], []

    monkeypatch.setattr(number_check_service, "_set_llm_env", lambda _model: "test-model")
    monkeypatch.setattr(
        number_check_service,
        "_load_latest_main_module",
        lambda: types.SimpleNamespace(run=fake_run),
    )

    task_id = "docx-output-init"
    number_check_service._init_task_progress(task_id)
    result = number_check_service._run_latest_number_check_sync(
        task_id=task_id,
        mode=number_check_service.NUMBER_CHECK_MODE_DIRECT,
        alignment_path=None,
        source_path=None,
        target_path=target_path,
        source_hf_path=None,
        output_dir=output_dir,
        gemini_route="openrouter",
        model_name="test-model",
        alignment_filename=None,
        source_filename=None,
        target_filename="translated.docx",
    )

    revised_path = captured["revised_path"]
    assert result["corrected_docx"].endswith(revised_path.name)
    assert revised_path.read_bytes() == target_content


def _run_direct(tmp_path, monkeypatch, source_name, target_name, fake_run):
    source_path, target_path = tmp_path / source_name, tmp_path / target_name
    source_path.write_bytes(b"src")
    target_path.write_bytes(b"tgt")
    output_dir = tmp_path / "outputs"
    output_dir.mkdir()

    monkeypatch.setattr(number_check_service, "_set_llm_env", lambda _model: "test-model")
    monkeypatch.setattr(
        number_check_service,
        "_load_latest_main_module",
        lambda: types.SimpleNamespace(run=fake_run),
    )
    task_id = f"direct-{source_name}"
    number_check_service._init_task_progress(task_id)
    return number_check_service._run_latest_number_check_sync(
        task_id=task_id,
        mode=number_check_service.NUMBER_CHECK_MODE_DIRECT,
        alignment_path=None,
        source_path=source_path,
        target_path=target_path,
        source_hf_path=None,
        output_dir=output_dir,
        gemini_route="openrouter",
        model_name="test-model",
        alignment_filename=None,
        source_filename=source_name,
        target_filename=target_name,
    )


def test_direct_docx_pair_uses_legacy_mode_and_handles_dict_result(tmp_path, monkeypatch):
    captured = {}

    def fake_run(**kwargs):
        captured.update(kwargs)
        return {"body": [{"错误编号": "1"}, {"错误编号": "2"}], "header": [{"错误编号": "1"}], "footer": []}

    result = _run_direct(tmp_path, monkeypatch, "source.docx", "target.docx", fake_run)

    assert captured["use_legacy_mode"] is True
    assert captured["bilingual_mode"] is False
    assert result["stats"] == {"total_issues": 3, "body_issues": 2, "header_issues": 1, "footer_issues": 0}


@pytest.mark.parametrize("ext", [".xlsx", ".pptx", ".pdf"])
def test_direct_non_docx_pair_also_uses_legacy_mode(tmp_path, monkeypatch, ext):
    captured = {}

    def fake_run(**kwargs):
        captured.update(kwargs)
        return {"body": [], "header": [], "footer": []}

    result = _run_direct(tmp_path, monkeypatch, f"source{ext}", f"target{ext}", fake_run)

    assert captured["use_legacy_mode"] is True
    assert captured["bilingual_mode"] is False
    assert result["stats"]["total_issues"] == 0
