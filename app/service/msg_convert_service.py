from __future__ import annotations

import base64
import hashlib
import html
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional
from urllib.parse import unquote, unquote_to_bytes, urlsplit

import extract_msg
from bs4 import BeautifulSoup, Comment, UnicodeDammit
from RTFDE.deencapsulate import DeEncapsulator

from app.core.config import settings
from app.core.file_naming import build_user_visible_filename
from app.service.libreoffice_service import (
    convert_docx_to_pdf_via_libreoffice,
    convert_to_docx_via_libreoffice,
)


MSG_CONVERT_ALLOWED_EXTENSIONS = {".msg"}
MSG_CONVERT_OUTPUT_FORMATS = {"word", "pdf", "both"}
MSG_CONVERT_DEFAULT_OUTPUT_FORMAT = "word"
MSG_CONVERT_MAX_FILES = 50
OLE_COMPOUND_FILE_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
MAX_INLINE_IMAGE_BYTES = 25 * 1024 * 1024

_DANGEROUS_TAGS = {
    "script",
    "iframe",
    "frame",
    "frameset",
    "object",
    "embed",
    "applet",
    "form",
    "input",
    "button",
    "textarea",
    "select",
    "option",
    "video",
    "audio",
    "source",
    "track",
    "canvas",
    "svg",
    "math",
    "foreignobject",
    "base",
}
_SAFE_LINK_SCHEMES = {"", "http", "https", "mailto", "tel"}
_BLOCKED_IMAGE_SCHEMES = {"http", "https", "ftp", "file", "javascript"}
_IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
    ".ico",
    ".emf",
    ".wmf",
}
_MIME_EXTENSION_MAP = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/webp": ".webp",
    "image/tiff": ".tiff",
    "image/x-icon": ".ico",
    "image/vnd.microsoft.icon": ".ico",
    "image/emf": ".emf",
    "image/x-emf": ".emf",
    "image/wmf": ".wmf",
    "image/x-wmf": ".wmf",
}
_CSS_URL_RE = re.compile(r"url\(\s*(['\"]?)(.*?)\1\s*\)", re.IGNORECASE | re.DOTALL)
_CSS_IMPORT_RE = re.compile(r"@import\s+(?:url\([^)]*\)|['\"][^'\"]*['\"])[^;]*;?", re.IGNORECASE)
_CSS_DANGEROUS_RE = re.compile(r"(?:expression\s*\(|javascript\s*:|vbscript\s*:|behavior\s*:|-moz-binding\s*:)", re.IGNORECASE)


@dataclass
class MessageBody:
    html: str
    body_format: str
    warnings: list[str]


@dataclass
class AttachmentImage:
    attachment: Any
    keys: set[str]
    filename: str
    mimetype: str
    inline_marked: bool


def normalize_msg_output_format(value: str | None) -> str:
    normalized = str(value or MSG_CONVERT_DEFAULT_OUTPUT_FORMAT).strip().lower()
    if normalized not in MSG_CONVERT_OUTPUT_FORMATS:
        allowed = "、".join(("word", "pdf", "both"))
        raise ValueError(f"output_format 必须是 {allowed} 之一")
    return normalized


def get_msg_convert_config() -> dict[str, Any]:
    return {
        "allowed_extensions": sorted(MSG_CONVERT_ALLOWED_EXTENSIONS),
        "max_files": MSG_CONVERT_MAX_FILES,
        "upload_max_mb": max(1, int(settings.MSG_CONVERT_UPLOAD_MAX_MB or 95)),
        "output_formats": {
            "word": {"label": "Word", "extensions": [".docx"]},
            "pdf": {"label": "PDF", "extensions": [".pdf"]},
            "both": {"label": "Word + PDF", "extensions": [".docx", ".pdf"]},
        },
        "default_output_format": MSG_CONVERT_DEFAULT_OUTPUT_FORMAT,
        "uses_ai": False,
        "downloads_external_images": False,
    }


def validate_msg_file(
    input_path: str | Path,
    original_filename: str | None = None,
    *,
    max_bytes: int | None = None,
) -> Path:
    path = Path(input_path)
    filename = original_filename or path.name
    if Path(filename).suffix.lower() not in MSG_CONVERT_ALLOWED_EXTENSIONS:
        raise ValueError("仅支持 .msg 文件")
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(f"MSG 文件不存在：{path}")
    size = path.stat().st_size
    if size <= 0:
        raise ValueError("MSG 文件为空")
    if max_bytes is not None and size > max_bytes:
        limit_mb = max_bytes / (1024 * 1024)
        raise ValueError(f"MSG 文件超过 {limit_mb:g} MB 限制")
    with path.open("rb") as handle:
        signature = handle.read(len(OLE_COMPOUND_FILE_SIGNATURE))
    if signature != OLE_COMPOUND_FILE_SIGNATURE:
        raise ValueError("文件不是有效的 Outlook MSG（OLE 文件头不正确）")
    return path


def _decode_html_bytes(value: bytes | str | None) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    decoded = UnicodeDammit(value, is_html=True).unicode_markup
    return decoded or value.decode("utf-8", errors="replace")


def _plain_text_to_html(value: str | bytes | None) -> str:
    if isinstance(value, bytes):
        value = value.decode("utf-8", errors="replace")
    text = str(value or "")
    return f'<div class="plain-text-body">{html.escape(text)}</div>'


def deencapsulate_rtf_body(rtf_body: bytes | None) -> tuple[str, str] | None:
    """返回 RTF 中封装的正文及其内容类型。"""
    if not rtf_body:
        return None
    deencapsulator = DeEncapsulator(rtf_body)
    deencapsulator.deencapsulate()
    content_type = str(getattr(deencapsulator, "content_type", "") or "").lower()
    if content_type == "html" and getattr(deencapsulator, "html", None):
        return _decode_html_bytes(deencapsulator.html), "html"
    text = getattr(deencapsulator, "text", None)
    if text is None:
        text = getattr(deencapsulator, "content", None)
    if text is not None:
        return _plain_text_to_html(text), "plain"
    return None


def select_message_body(
    raw_html: bytes | str | None,
    rtf_body: bytes | None,
    plain_body: bytes | str | None,
) -> MessageBody:
    warnings: list[str] = []
    if raw_html:
        return MessageBody(_decode_html_bytes(raw_html), "html", warnings)

    if rtf_body:
        try:
            deencapsulated = deencapsulate_rtf_body(rtf_body)
        except Exception as exc:
            warnings.append(f"RTF 正文解封装失败，已回退到纯文本：{exc}")
        else:
            if deencapsulated and deencapsulated[0].strip():
                return MessageBody(deencapsulated[0], "rtf", warnings)
            warnings.append("RTF 正文为空，已回退到纯文本")

    plain_has_content = bool(plain_body.strip()) if isinstance(plain_body, bytes) else bool(str(plain_body or "").strip())
    if plain_has_content:
        return MessageBody(_plain_text_to_html(plain_body), "plain", warnings)

    warnings.append("邮件正文为空")
    return MessageBody('<p class="empty-body">（邮件正文为空）</p>', "plain", warnings)


def _safe_get(obj: Any, name: str, default: Any = None) -> Any:
    try:
        value = getattr(obj, name)
    except Exception:
        return default
    return default if value is None else value


def _normalize_reference(value: str | None) -> set[str]:
    if not value:
        return set()
    raw = html.unescape(str(value)).strip().strip("'\"")
    if not raw:
        return set()
    decoded = unquote(raw).strip()
    values = {raw, decoded}
    for item in tuple(values):
        stripped = item.strip().strip("<>")
        values.add(stripped)
        if stripped.lower().startswith("cid:"):
            values.add(stripped[4:].strip().strip("<>"))
        path = urlsplit(stripped).path
        if path:
            values.add(Path(path.replace("\\", "/")).name)
        values.add(stripped.lstrip("./"))
    return {item.casefold() for item in values if item}


def _attachment_content_location(attachment: Any) -> str:
    direct = _safe_get(attachment, "contentLocation", "")
    if direct:
        return str(direct)
    getter = _safe_get(attachment, "getStringStream")
    if callable(getter):
        try:
            return str(getter("__substg1.0_3713") or "")
        except Exception:
            return ""
    return ""


def _detect_image_extension(data: bytes, mimetype: str = "", filename: str = "") -> str | None:
    mime = str(mimetype or "").split(";", 1)[0].strip().lower()
    if mime == "image/svg+xml":
        return None
    if mime in _MIME_EXTENSION_MAP:
        return _MIME_EXTENSION_MAP[mime]
    suffix = Path(filename or "").suffix.lower()
    if suffix in _IMAGE_EXTENSIONS:
        return ".jpg" if suffix == ".jpeg" else suffix
    signatures = (
        (b"\x89PNG\r\n\x1a\n", ".png"),
        (b"\xff\xd8\xff", ".jpg"),
        (b"GIF87a", ".gif"),
        (b"GIF89a", ".gif"),
        (b"BM", ".bmp"),
        (b"II*\x00", ".tiff"),
        (b"MM\x00*", ".tiff"),
        (b"\x00\x00\x01\x00", ".ico"),
    )
    for signature, extension in signatures:
        if data.startswith(signature):
            return extension
    if data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        return ".webp"
    return None


def collect_attachment_images(message: Any) -> list[AttachmentImage]:
    images: list[AttachmentImage] = []
    for attachment in _safe_get(message, "attachments", []) or []:
        attachment_type = str(_safe_get(attachment, "type", "")).lower()
        if "msg" in attachment_type and "data" not in attachment_type:
            continue
        cid = str(_safe_get(attachment, "cid", "") or _safe_get(attachment, "contentId", "") or "")
        location = _attachment_content_location(attachment)
        filename = str(
            _safe_get(attachment, "longFilename", "")
            or _safe_get(attachment, "shortFilename", "")
            or _safe_get(attachment, "name", "")
            or "inline-image"
        )
        mimetype = str(_safe_get(attachment, "mimetype", "") or "")
        suffix = Path(filename).suffix.lower()
        is_known_image = mimetype.lower().startswith("image/") or suffix in _IMAGE_EXTENSIONS
        rendering_position = _safe_get(attachment, "renderingPosition")
        inline_marked = bool(cid or location or _safe_get(attachment, "hidden", False))
        if isinstance(rendering_position, int) and rendering_position != 0xFFFFFFFF:
            inline_marked = True
        if not is_known_image and not inline_marked:
            continue

        keys: set[str] = set()
        for reference in (cid, location, filename):
            keys.update(_normalize_reference(reference))
        images.append(
            AttachmentImage(
                attachment=attachment,
                keys=keys,
                filename=Path(filename).name,
                mimetype=mimetype,
                inline_marked=inline_marked,
            )
        )
    return images


class InlineImageResolver:
    def __init__(self, assets_dir: Path, attachment_images: list[AttachmentImage], warnings: list[str]):
        self.assets_dir = assets_dir
        self.attachment_images = attachment_images
        self.warnings = warnings
        self._key_map: dict[str, AttachmentImage] = {}
        self._saved_attachments: dict[int, str] = {}
        self._saved_hashes: dict[str, str] = {}
        self._used_paths: set[str] = set()
        for item in attachment_images:
            for key in item.keys:
                self._key_map.setdefault(key, item)

    @property
    def inline_count(self) -> int:
        return len(self._used_paths)

    def _write_image(self, data: bytes, mimetype: str = "", filename: str = "") -> str | None:
        if not data:
            return None
        if len(data) > MAX_INLINE_IMAGE_BYTES:
            self.warnings.append(f"内嵌图片 {filename or '未命名图片'} 超过 25 MB，已跳过")
            return None
        extension = _detect_image_extension(data, mimetype, filename)
        if not extension:
            self.warnings.append(f"无法识别内嵌图片格式：{filename or '未命名图片'}")
            return None
        digest = hashlib.sha256(data).hexdigest()
        if digest in self._saved_hashes:
            path = self._saved_hashes[digest]
            self._used_paths.add(path)
            return path
        self.assets_dir.mkdir(parents=True, exist_ok=True)
        asset_name = f"inline_{len(self._saved_hashes) + 1:03d}_{digest[:8]}{extension}"
        asset_path = self.assets_dir / asset_name
        asset_path.write_bytes(data)
        relative_path = f"assets/{asset_name}"
        self._saved_hashes[digest] = relative_path
        self._used_paths.add(relative_path)
        return relative_path

    def _save_attachment(self, item: AttachmentImage) -> str | None:
        attachment_id = id(item.attachment)
        if attachment_id in self._saved_attachments:
            path = self._saved_attachments[attachment_id]
            self._used_paths.add(path)
            return path
        data = _safe_get(item.attachment, "data")
        if not isinstance(data, (bytes, bytearray, memoryview)):
            self.warnings.append(f"内嵌图片无法读取：{item.filename}")
            return None
        path = self._write_image(bytes(data), item.mimetype, item.filename)
        if path:
            self._saved_attachments[attachment_id] = path
        return path

    def resolve(self, reference: str | None) -> str | None:
        if not reference:
            return None
        value = html.unescape(str(reference)).strip().strip("'\"")
        if value.lower().startswith("data:"):
            return self._resolve_data_uri(value)
        for key in _normalize_reference(value):
            item = self._key_map.get(key)
            if item:
                return self._save_attachment(item)
        return None

    def appendable_inline_images(self) -> list[tuple[str, str]]:
        resolved: list[tuple[str, str]] = []
        for item in self.attachment_images:
            if not item.inline_marked:
                continue
            path = self._save_attachment(item)
            if path and all(existing_path != path for existing_path, _ in resolved):
                resolved.append((path, item.filename))
        return resolved

    def _resolve_data_uri(self, value: str) -> str | None:
        try:
            header, payload = value.split(",", 1)
        except ValueError:
            self.warnings.append("发现无效的 data URI 图片，已跳过")
            return None
        media_info = header[5:]
        mimetype = media_info.split(";", 1)[0].lower()
        if mimetype not in _MIME_EXTENSION_MAP:
            self.warnings.append(f"不支持的 data URI 图片类型：{mimetype or '未知'}")
            return None
        try:
            if ";base64" in media_info.lower():
                data = base64.b64decode(payload, validate=True)
            else:
                data = unquote_to_bytes(payload)
        except Exception:
            self.warnings.append("data URI 图片解码失败，已跳过")
            return None
        return self._write_image(data, mimetype, "data-uri")


def _replace_unavailable_image(tag: Any, *, external_url: str | None = None) -> None:
    alt = str(tag.get("alt") or tag.get("title") or "图片")
    owner = tag if isinstance(tag, BeautifulSoup) else tag.find_parent()
    while owner is not None and not isinstance(owner, BeautifulSoup):
        owner = getattr(owner, "parent", None)
    if owner is None:
        tag.replace_with(f"[{alt}]")
        return
    span = owner.new_tag("span")
    span["class"] = "blocked-image"
    if external_url:
        anchor = owner.new_tag("a", href=external_url)
        anchor.string = f"[外部图片：{alt}]"
        span.append(anchor)
    else:
        span.string = f"[{alt}]"
    tag.replace_with(span)


def _sanitize_css(css: str, resolver: InlineImageResolver) -> str:
    value = _CSS_IMPORT_RE.sub("", css or "")
    value = _CSS_DANGEROUS_RE.sub("blocked(", value)

    def replace_url(match: re.Match[str]) -> str:
        raw_url = html.unescape(match.group(2) or "").strip()
        local_path = resolver.resolve(raw_url)
        if local_path:
            return f'url("{local_path}")'
        return "none"

    return _CSS_URL_RE.sub(replace_url, value)


def sanitize_email_html(
    body_html: str,
    resolver: InlineImageResolver,
) -> tuple[str, str]:
    soup = BeautifulSoup(body_html or "", "html.parser")
    for comment in soup.find_all(string=lambda value: isinstance(value, Comment)):
        comment.extract()
    for tag in list(soup.find_all(_DANGEROUS_TAGS)):
        tag.decompose()
    for tag in list(soup.find_all("link")):
        tag.decompose()
    for tag in list(soup.find_all("meta")):
        http_equiv = str(tag.get("http-equiv") or "").lower()
        if http_equiv in {"refresh", "content-security-policy"}:
            tag.decompose()

    for tag in soup.find_all(True):
        for attr_name in list(tag.attrs):
            lowered = attr_name.lower()
            if lowered.startswith("on") or lowered in {
                "srcset",
                "formaction",
                "poster",
                "background",
                "lowsrc",
                "dynsrc",
                "o:href",
            }:
                del tag.attrs[attr_name]
        if tag.has_attr("style"):
            tag["style"] = _sanitize_css(str(tag.get("style") or ""), resolver)
        if tag.name == "a" and tag.has_attr("href"):
            href = str(tag.get("href") or "").strip()
            scheme = urlsplit(href).scheme.lower()
            if scheme not in _SAFE_LINK_SCHEMES:
                del tag.attrs["href"]

    for style_tag in soup.find_all("style"):
        safe_css = _sanitize_css(style_tag.get_text("", strip=False), resolver)
        style_tag.clear()
        style_tag.append(safe_css)

    image_tags = list(soup.find_all(["img", "v:imagedata"]))
    for tag in image_tags:
        source = str(tag.get("src") or tag.get("href") or tag.get("xlink:href") or "").strip()
        local_path = resolver.resolve(source)
        if local_path:
            if tag.name != "img":
                replacement = soup.new_tag("img")
                for attr in ("alt", "title", "width", "height", "style"):
                    if tag.get(attr) is not None:
                        replacement[attr] = tag.get(attr)
                tag.replace_with(replacement)
                tag = replacement
            tag["src"] = local_path
            tag.attrs.pop("href", None)
            tag.attrs.pop("xlink:href", None)
            continue

        scheme = urlsplit(source).scheme.lower()
        is_external = source.startswith("//") or scheme in _BLOCKED_IMAGE_SCHEMES
        if is_external:
            resolver.warnings.append("外部图片未下载，已保留替代文本或链接")
            _replace_unavailable_image(tag, external_url=source if scheme in {"http", "https"} else None)
        else:
            resolver.warnings.append(f"正文图片未在 MSG 内找到：{source or '无地址'}")
            _replace_unavailable_image(tag)

    styles: list[str] = []
    for style_tag in list(soup.find_all("style")):
        styles.append(str(style_tag))
        style_tag.extract()
    container = soup.body if soup.body else soup
    fragment = "".join(str(child) for child in container.contents)
    return fragment, "\n".join(styles)


def _format_date(value: Any) -> str:
    if isinstance(value, datetime):
        timezone = value.strftime(" %z") if value.tzinfo else ""
        return f"{value:%Y-%m-%d %H:%M:%S}{timezone}"
    return str(value or "")


def _metadata_table(message: Any, subject: str) -> str:
    rows = [
        ("主题", subject),
        ("发件人", _safe_get(message, "sender", "")),
        ("收件人", _safe_get(message, "to", "")),
        ("抄送", _safe_get(message, "cc", "")),
        ("密送", _safe_get(message, "bcc", "")),
        ("日期", _format_date(_safe_get(message, "date", ""))),
    ]
    cells = []
    for label, value in rows:
        text = str(value or "").strip()
        if not text and label not in {"主题", "日期"}:
            continue
        cells.append(f"<tr><th>{html.escape(label)}</th><td>{html.escape(text or '—')}</td></tr>")
    return '<table class="message-info">' + "".join(cells) + "</table>"


def build_safe_message_html(message: Any, assets_dir: str | Path) -> dict[str, Any]:
    warnings: list[str] = []
    raw_html = None
    stream_getter = _safe_get(message, "getStream")
    if callable(stream_getter):
        try:
            raw_html = stream_getter("__substg1.0_10130102")
        except Exception as exc:
            warnings.append(f"HTML 正文读取失败：{exc}")
    rtf_body = _safe_get(message, "rtfBody")
    plain_body = None
    string_getter = _safe_get(message, "getStringStream")
    if callable(string_getter):
        try:
            plain_body = string_getter("__substg1.0_1000")
        except Exception as exc:
            warnings.append(f"纯文本正文读取失败：{exc}")
    if plain_body is None:
        plain_body = _safe_get(message, "body")

    selected = select_message_body(raw_html, rtf_body, plain_body)
    warnings.extend(selected.warnings)
    attachment_images = collect_attachment_images(message)
    resolver = InlineImageResolver(Path(assets_dir), attachment_images, warnings)
    safe_body, original_styles = sanitize_email_html(selected.html, resolver)

    appended_images_html = ""
    if resolver.inline_count == 0:
        appended = resolver.appendable_inline_images()
        if appended:
            image_items = "".join(
                f'<figure><img src="{html.escape(path, quote=True)}" alt="{html.escape(name, quote=True)}"><figcaption>{html.escape(name)}</figcaption></figure>'
                for path, name in appended
            )
            appended_images_html = (
                '<section class="inline-image-gallery"><h2>邮件内嵌图片</h2>'
                f'<div class="inline-image-grid">{image_items}</div></section>'
            )
            warnings.append("正文中未找到图片位置，已将 MSG 标记的内嵌图片附加到正文末尾")

    subject = str(_safe_get(message, "subject", "") or "").strip() or "（无主题）"
    document = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>{html.escape(subject)}</title>
<style>
@page {{ size: A4; margin: 18mm; }}
body {{ font-family: "Noto Sans CJK SC", "Microsoft YaHei", Arial, sans-serif; color: #1f2937; font-size: 10.5pt; line-height: 1.55; }}
.message-info {{ width: 100%; border-collapse: collapse; margin: 0 0 20px; table-layout: fixed; }}
.message-info th, .message-info td {{ border: 1px solid #cbd5e1; padding: 7px 9px; vertical-align: top; overflow-wrap: anywhere; }}
.message-info th {{ width: 68px; background: #f1f5f9; text-align: left; white-space: nowrap; }}
.message-body {{ overflow-wrap: anywhere; }}
.message-body img, .inline-image-gallery img {{ max-width: 100% !important; height: auto !important; }}
.plain-text-body {{ white-space: pre-wrap; font-family: "Noto Sans CJK SC", "Microsoft YaHei", sans-serif; }}
.blocked-image {{ color: #64748b; font-style: italic; }}
.inline-image-gallery {{ margin-top: 24px; border-top: 1px solid #cbd5e1; padding-top: 14px; }}
.inline-image-gallery h2 {{ font-size: 12pt; }}
.inline-image-grid figure {{ margin: 10px 0; page-break-inside: avoid; }}
.inline-image-grid figcaption {{ color: #64748b; font-size: 9pt; }}
table {{ max-width: 100%; }}
a {{ color: #0369a1; text-decoration: underline; }}
</style>
{original_styles}
</head>
<body>
{_metadata_table(message, subject)}
<main class="message-body">{safe_body}</main>
{appended_images_html}
</body>
</html>"""
    return {
        "html": document,
        "subject": subject,
        "body_format": selected.body_format,
        "inline_image_count": resolver.inline_count,
        "warnings": list(dict.fromkeys(warnings)),
    }


def convert_msg_to_documents(
    *,
    input_path: str | Path,
    original_filename: str,
    display_no: str,
    output_format: str = MSG_CONVERT_DEFAULT_OUTPUT_FORMAT,
) -> dict[str, Any]:
    normalized_format = normalize_msg_output_format(output_format)
    max_bytes = max(1, int(settings.MSG_CONVERT_UPLOAD_MAX_MB or 95)) * 1024 * 1024
    validated_path = validate_msg_file(input_path, original_filename, max_bytes=max_bytes)

    output_dir = Path(settings.OUTPUT_DIR) / "msg_convert" / display_no
    output_dir.mkdir(parents=True, exist_ok=True)
    assets_dir = output_dir / "assets"
    html_path = output_dir / "message_source.html"
    visible_docx_name = build_user_visible_filename(original_filename, ext=".docx")
    visible_pdf_name = build_user_visible_filename(original_filename, ext=".pdf")
    docx_path = output_dir / visible_docx_name
    pdf_path = output_dir / visible_pdf_name

    message = None
    completed = False
    prepared: dict[str, Any]
    try:
        try:
            message = extract_msg.openMsg(str(validated_path), strict=True)
            prepared = build_safe_message_html(message, assets_dir)
        except Exception as exc:
            raise ValueError(f"MSG 文件解析失败：{exc}") from exc
        finally:
            if message is not None:
                try:
                    message.close()
                except Exception:
                    pass

        html_path.write_text(prepared["html"], encoding="utf-8")
        convert_to_docx_via_libreoffice(html_path, docx_path)
        if normalized_format in {"pdf", "both"}:
            convert_docx_to_pdf_via_libreoffice(docx_path, pdf_path)

        published_docx = docx_path if normalized_format in {"word", "both"} else None
        published_pdf = pdf_path if normalized_format in {"pdf", "both"} else None
        if normalized_format == "pdf":
            docx_path.unlink(missing_ok=True)

        def portable(path: Path | None) -> str | None:
            return str(path).replace("\\", "/") if path else None

        result = {
            "filename": original_filename,
            "subject": prepared["subject"],
            "body_format": prepared["body_format"],
            "inline_image_count": prepared["inline_image_count"],
            "output_format": normalized_format,
            "output_docx": portable(published_docx),
            "output_pdf": portable(published_pdf),
            "warnings": prepared["warnings"],
        }
        completed = True
        return result
    finally:
        html_path.unlink(missing_ok=True)
        shutil.rmtree(assets_dir, ignore_errors=True)
        if normalized_format == "pdf":
            docx_path.unlink(missing_ok=True)
        if not completed:
            docx_path.unlink(missing_ok=True)
            pdf_path.unlink(missing_ok=True)


async def execute_msg_convert_task(
    *,
    task_id: str,
    display_no: str,
    input_path: str,
    original_filename: str,
    output_format: str,
    progress_callback: Optional[Callable[[int, str], Any]] = None,
    executor=None,
) -> dict[str, Any]:
    import asyncio

    if progress_callback:
        await progress_callback(10, "正在解析 MSG 邮件")
    loop = asyncio.get_running_loop()
    result = await loop.run_in_executor(
        executor,
        lambda: convert_msg_to_documents(
            input_path=input_path,
            original_filename=original_filename,
            display_no=display_no,
            output_format=output_format,
        ),
    )
    if progress_callback:
        await progress_callback(95, "文档转换完成，正在登记输出")
    return result
