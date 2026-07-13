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
