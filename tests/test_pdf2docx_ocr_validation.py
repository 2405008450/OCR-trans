import base64
import io

import fitz
import pytest
from PIL import Image

import pdf2docx as converter


EMPTY_HTML = '<!DOCTYPE html><html><head><title>Document</title></head><body></body></html>'


@pytest.fixture
def dark_image():
    output = io.BytesIO()
    Image.new('RGB', (160, 160), '#354050').save(output, format='PNG')
    return output.getvalue()


@pytest.fixture
def single_attempt(monkeypatch):
    monkeypatch.setattr(converter, '_build_ocr_attempt_plan', lambda *args: [
        {'route': 'openrouter', 'model': 'test-model', 'retries': 1}
    ])
    monkeypatch.setattr(converter.time, 'sleep', lambda _: None)


def test_illustration_can_return_explicit_empty_html(monkeypatch, dark_image, single_attempt):
    assert converter._image_has_visible_text_like_content(dark_image)
    calls = []

    def generate(**kwargs):
        calls.append(kwargs)
        return EMPTY_HTML

    monkeypatch.setattr(converter, 'generate_vision_html', generate)
    assert converter._ocr_single_image(base64.b64encode(dark_image), 'image/png', 'test-model') == EMPTY_HTML
    assert len(calls) == 1


@pytest.mark.parametrize('response', ['', 'blank', '<html><body>', '<html><body></body>'])
def test_incomplete_results_still_fail(monkeypatch, dark_image, single_attempt, response):
    monkeypatch.setattr(converter, 'generate_vision_html', lambda **kwargs: response)
    with pytest.raises(RuntimeError, match='OCR 失败'):
        converter._ocr_single_image(base64.b64encode(dark_image), 'image/png', 'test-model')


def test_error_message_distinguishes_empty_from_truncated():
    error = converter.OCRIncompleteResultError('OCR 未返回明确的识别结果')
    assert converter._ocr_exception_message(error) == str(error)


def test_illustration_does_not_stop_later_pdf_pages(tmp_path, monkeypatch, dark_image, single_attempt):
    path = tmp_path / 'illustration.pdf'
    with fitz.open() as document:
        for _ in range(3):
            page = document.new_page()
            page.insert_image(page.rect, stream=dark_image)
        document.save(path)
    responses = iter(['<html><body><p>第一页</p></body></html>', EMPTY_HTML,
                      '<html><body><p>第三页</p></body></html>'])
    monkeypatch.setattr(converter, 'generate_vision_html', lambda **kwargs: next(responses))
    result = converter.ocr_file(str(path), return_metadata=True)
    assert result['total_pages'] == 3
    assert result['blank_pages'] == [2]
    assert result['failed_pages'] == []
    assert result['text'].count('<page_break/>') == 2
    assert '第三页' in result['text']
