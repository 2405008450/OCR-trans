from pathlib import Path

import re
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from app.controller import task
from app.core.config import settings
from app.db.init_db import init_db
from app.service.task_queue_service import ocr_task_queue

BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "static"
UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUTS_DIR = BASE_DIR / "outputs"
TEMP_IMAGES_DIR = BASE_DIR / "temp_images"
HTML_CACHE_HEADERS = {
    "Cache-Control": "no-cache, no-store, must-revalidate",
    "Pragma": "no-cache",
    "Expires": "0",
}
LOCAL_STATIC_ASSET_PATTERN = re.compile(r'(?P<quote>["\'])(?P<url>/static/[^"\']+)(?P=quote)')

app = FastAPI(
    title="图片 OCR 翻译系统",
    description="AI 驱动的文档识别、翻译与处理平台",
    version="1.0.0",
)

allowed_origins = settings.ALLOWED_ORIGINS.split(",") if settings.ALLOWED_ORIGINS != "*" else ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
if UPLOADS_DIR.exists():
    app.mount("/uploads", StaticFiles(directory=str(UPLOADS_DIR)), name="uploads")
if OUTPUTS_DIR.exists():
    app.mount("/outputs", StaticFiles(directory=str(OUTPUTS_DIR)), name="outputs")
if TEMP_IMAGES_DIR.exists():
    app.mount("/temp_images", StaticFiles(directory=str(TEMP_IMAGES_DIR)), name="temp_images")

app.include_router(task.router)


@app.on_event("startup")
async def startup_event():
    init_db()
    await ocr_task_queue.start()


@app.on_event("shutdown")
async def shutdown_event():
    await ocr_task_queue.stop()


def _read_page(filename: str) -> str:
    page_file = STATIC_DIR / filename
    if page_file.exists():
        return page_file.read_text(encoding="utf-8")
    return f"<h1>错误：找不到 {filename}</h1><p>路径: {page_file}</p>"


def _compute_app_build_id() -> str:
    latest_mtime = 0.0
    if STATIC_DIR.exists():
        for path in STATIC_DIR.rglob("*"):
            if path.is_file():
                latest_mtime = max(latest_mtime, path.stat().st_mtime)
    return str(int(latest_mtime)) if latest_mtime else "dev"


APP_BUILD_ID = _compute_app_build_id()


def _version_static_url(url: str) -> str:
    parts = urlsplit(url)
    query_items = dict(parse_qsl(parts.query, keep_blank_values=True))
    query_items["v"] = APP_BUILD_ID
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query_items), parts.fragment))


def _inject_app_build_meta(html: str) -> str:
    meta_tag = f'<meta name="app-build-id" content="{APP_BUILD_ID}">'
    if meta_tag in html:
        return html
    return html.replace("<head>", f"<head>\n    {meta_tag}", 1)


def _inject_static_asset_version(html: str) -> str:
    def _replace(match: re.Match[str]) -> str:
        quote = match.group("quote")
        url = match.group("url")
        return f"{quote}{_version_static_url(url)}{quote}"

    return LOCAL_STATIC_ASSET_PATTERN.sub(_replace, html)


def _render_page(filename: str) -> HTMLResponse:
    content = _read_page(filename)
    content = _inject_app_build_meta(content)
    content = _inject_static_asset_version(content)
    return HTMLResponse(content=content, headers=HTML_CACHE_HEADERS)


@app.get("/", response_class=HTMLResponse)
async def root():
    return _render_page("nav.html")


@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard_page():
    return _render_page("dashboard.html")


@app.get("/ocr", response_class=HTMLResponse)
async def ocr_page():
    return _render_page("index.html")


@app.get("/certificate-translation", response_class=HTMLResponse)
async def certificate_translation_page():
    return _render_page("certificate_translation.html")


@app.get("/number-check", response_class=HTMLResponse)
async def number_check_page():
    return _render_page("number_check.html")


@app.get("/alignment", response_class=HTMLResponse)
async def alignment_page():
    return _render_page("alignment.html")


@app.get("/drivers-license", response_class=HTMLResponse)
async def drivers_license_page():
    return _render_page("drivers_license.html")


@app.get("/doc-translate", response_class=HTMLResponse)
async def doc_translate_page():
    return _render_page("doc_translate.html")


@app.get("/business-licence/embed", response_class=HTMLResponse)
async def business_licence_embed_page():
    return _render_page("business_licence.html")


@app.get("/business-licence")
async def business_licence_page():
    return RedirectResponse(url="/certificate-translation?tab=business-panel", status_code=307)


@app.get("/zhongfanyi", response_class=HTMLResponse)
async def zhongfanyi_page():
    return _render_page("zhongfanyi.html")


@app.get("/pdf2docx", response_class=HTMLResponse)
async def pdf2docx_page():
    return _render_page("pdf2docx.html")


@app.get("/healthz")
async def healthz():
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "app.main:app",
        host=settings.HOST,
        port=settings.PORT,
        reload=settings.DEBUG_ENABLED,
    )
