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


@pytest.mark.parametrize('encoding', ['gbk', 'ascii', 'utf-8'])
def test_unicode_ocr_result_does_not_retry_for_console_encoding(monkeypatch, dark_image, encoding):
    response = '<html><body><p>中文・🙂</p></body></html>'
    calls = []

    def generate(**kwargs):
        calls.append(kwargs)
        return response

    monkeypatch.setattr(converter, 'generate_vision_html', generate)
    monkeypatch.setattr(converter, '_build_ocr_attempt_plan', lambda *args: [
        {'route': 'openrouter', 'model': 'test-model', 'retries': 3},
        {'route': 'aistudio', 'model': 'fallback-model', 'retries': 1},
    ])

    def unexpected_sleep(_):
        pytest.fail('Console encoding must not trigger an OCR retry')

    monkeypatch.setattr(converter.time, 'sleep', unexpected_sleep)
    buffer = io.BytesIO()
    with io.TextIOWrapper(buffer, encoding=encoding, errors='strict') as stdout:
        with monkeypatch.context() as console:
            console.setattr(converter.sys, 'stdout', stdout)
            result = converter._ocr_single_image(base64.b64encode(dark_image), 'image/png', 'test-model')
        logged = buffer.getvalue().decode(encoding)

    assert result == response
    assert len(calls) == 1
    assert logged == response.encode(encoding, errors='backslashreplace').decode(encoding)


def test_gbk_console_does_not_block_original_status_callback(monkeypatch):
    message = '↘ OCR 降级：中文・🙂'
    statuses = []
    buffer = io.BytesIO()
    with io.TextIOWrapper(buffer, encoding='gbk', errors='strict', newline='\n') as stdout:
        with monkeypatch.context() as console:
            console.setattr(converter.sys, 'stdout', stdout)
            converter._emit_ocr_status(message, statuses.append)
        logged = buffer.getvalue().decode('gbk')

    assert statuses == [message]
    assert logged == message.encode('gbk', errors='backslashreplace').decode('gbk') + '\n'


def test_gbk_console_allows_fallback_for_real_ocr_errors(monkeypatch, dark_image):
    response = '<html><body><p>中文・🙂</p></body></html>'
    calls = []
    statuses = []

    def generate(**kwargs):
        calls.append(kwargs)
        if len(calls) == 1:
            raise RuntimeError('请求失败・🙂')
        return response

    monkeypatch.setattr(converter, 'generate_vision_html', generate)
    monkeypatch.setattr(converter, '_build_ocr_attempt_plan', lambda *args: [
        {'route': 'openrouter', 'model': 'test-model', 'retries': 1},
        {'route': 'aistudio', 'model': 'fallback-model', 'retries': 1},
    ])
    with io.TextIOWrapper(io.BytesIO(), encoding='gbk', errors='strict') as stdout:
        with monkeypatch.context() as console:
            console.setattr(converter.sys, 'stdout', stdout)
            result = converter._ocr_single_image(
                base64.b64encode(dark_image), 'image/png', 'test-model', status_callback=statuses.append,
            )

    assert result == response
    assert [call['route'] for call in calls] == ['openrouter', 'aistudio']
    assert any('请求失败・🙂' in status for status in statuses)
    assert any(status.startswith('↘ OCR 降级') for status in statuses)
    assert statuses[-1].startswith('OCR 已恢复')
