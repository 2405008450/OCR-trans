import inspect
import sys
import types
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.service import number_check_service


def test_number_check_service_uses_v2_specialist_root():
    expected_root = Path(__file__).resolve().parents[1] / "专检" / "数检_程序-AIV2"

    assert number_check_service.NUMBER_CHECK_LATEST_ROOT == expected_root
    assert number_check_service.NUMBER_CHECK_MAIN_FILE == expected_root / "main.py"
    assert number_check_service.NUMBER_CHECK_MAIN_FILE.is_file()


def test_clear_specialist_module_cache_removes_modules_from_specialist_tree():
    module_name = "number_check_v2_cache_probe"
    module = types.ModuleType(module_name)
    module.__file__ = str(number_check_service.NUMBER_CHECK_LATEST_ROOT / "cache_probe.py")
    sys.modules[module_name] = module

    number_check_service._clear_specialist_module_cache()

    assert module_name not in sys.modules


def test_v2_main_keeps_system_integration_contract(monkeypatch):
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
    }.issubset(parameters)


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
