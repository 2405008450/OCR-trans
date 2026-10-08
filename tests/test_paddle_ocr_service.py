import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app.service import paddle_ocr_service as paddle
from app.service import word_count_service as service


def test_disabled_does_not_import_or_call_engine(monkeypatch):
    monkeypatch.setattr(paddle.settings, "PADDLEOCR_ENABLED", False)
    monkeypatch.setattr(paddle, "_get_engine", lambda _: pytest.fail("禁止加载"))
    with pytest.raises(RuntimeError, match="尚未启用"):
        paddle.extract_paddle_plain_text(file_path="missing.pdf")
    assert service.get_word_count_ocr_models()[paddle.MODEL_ID]["disabled"]


def test_paddle_selection_never_calls_llm(monkeypatch):
    monkeypatch.setattr(service, "extract_ocr_plain_text", lambda **_: pytest.fail("禁止调用LLM"))
    monkeypatch.setattr(service, "extract_paddle_plain_text", lambda **_: {"text": "sample"})
    assert service._extract_word_count_ocr(model=paddle.MODEL_ID, file_path="test.pdf")["text"] == "sample"


def test_cache_review_and_failure_retry(tmp_path, monkeypatch):
    import fitz
    source = tmp_path / "scan.pdf"
    with fitz.open() as doc:
        doc.new_page()
        doc.new_page()
        doc.save(source)
    monkeypatch.setattr(paddle.settings, "PADDLEOCR_ENABLED", True)
    monkeypatch.setattr(paddle.settings, "PADDLEOCR_CACHE_DIR", str(tmp_path / "cache"))
    monkeypatch.setattr(paddle.importlib.metadata, "version", lambda _: "test")
    monkeypatch.setattr(paddle, "_get_engine", lambda _: object())
    calls = []

    def recognize(page, engine, dpi):
        calls.append(page.number)
        if page.number == 1:
            raise RuntimeError("sample failure")
        return [{"text": "Material specification", "score": 0.7, "bbox": [0, 0, 10, 10]}]

    monkeypatch.setattr(paddle, "_recognize_page", recognize)
    result = paddle.extract_paddle_plain_text(file_path=str(source))
    assert result["failed_pages"] == [2]
    assert result["review_pages"] == [1]
    paddle.extract_paddle_plain_text(file_path=str(source))
    assert calls == [0, 1, 1]  # 成功页复用缓存，失败页重试。


def test_sparse_pdf_replaces_page_text_without_double_count(tmp_path, monkeypatch):
    import fitz
    source = tmp_path / "mixed.pdf"
    from PIL import Image
    image = tmp_path / "image.png"
    Image.new("RGB", (50, 50), "white").save(image)
    with fitz.open() as doc:
        page = doc.new_page()
        page.insert_image(page.rect, filename=str(image))
        page.insert_text((30, 50), "MATERIALS")
        doc.save(source)
    monkeypatch.setattr(service, "extract_paddle_plain_text", lambda **_: {
        "page_results": [{"page_number": 1, "text": "MATERIALS SPECIFICATIONS", "error": ""}],
        "processed_pages": [1], "failed_pages": [], "review_pages": [1]})
    content, metadata = service._extract_pdf_content_with_ocr(source, model=paddle.MODEL_ID, gemini_route="auto")
    assert [item.text for item in content.items] == ["MATERIALS SPECIFICATIONS"]
    assert service.count_words_word_like(content.items[0].text) == 2
    assert "复核页：1" in content.warning
    assert metadata["ocr_review_pages"] == [1]


def test_same_text_at_different_positions_is_not_deduplicated():
    assert paddle._same_region([0, 0, 100, 20], [1, 0, 101, 20])
    assert not paddle._same_region([0, 0, 100, 20], [0, 100, 100, 120])


def test_complete_line_replaces_multiple_high_confidence_tile_fragments():
    lines = [{"text":"Suspended", "score":0.99, "bbox":[0,0,30,10]},
             {"text":"ceilings make use of", "score":0.98, "bbox":[35,0,100,10]}]
    full = {"text":"Suspended ceilings make use of", "score":0.9, "bbox":[0,0,100,10]}
    paddle._merge_line(lines, full)
    assert lines == [full]
    paddle._merge_line(lines, {"text":"Suspended", "score":0.999, "bbox":[0,0,30,10]})
    assert lines == [full]


def test_normal_image_is_recognized_whole_without_tile_boundaries(tmp_path):
    import fitz
    from PIL import Image
    source = tmp_path / 'page.png'
    Image.new('RGB', (2000, 2500), 'white').save(source)
    calls = []
    class Engine:
        def predict(self, array):
            calls.append(array.shape)
            return []
    with fitz.open(source) as doc:
        paddle._recognize_page(doc[0], Engine(), 72)
    assert len(calls) == 1


def test_no_recognized_text_is_not_successful_zero_quote(tmp_path, monkeypatch):
    import fitz
    source = tmp_path / "empty.pdf"
    with fitz.open() as doc:
        doc.new_page()
        doc.save(source)
    monkeypatch.setattr(service, "extract_paddle_plain_text", lambda **_: {
        "page_results": [{"page_number": 1, "text": "", "error": ""}],
        "processed_pages": [1], "failed_pages": [], "review_pages": [1]})
    result, _, _ = service._count_single_file(file_path=source, root=tmp_path,
        converted_dir=tmp_path / "converted", max_bytes=1024*1024,
        ocr_enabled=True, ocr_model=paddle.MODEL_ID)
    assert result["status"] == service.STATUS_NEEDS_OCR
    assert "人工确认" in result["message"]
