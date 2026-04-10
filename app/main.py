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
REFRESH_HINT = "页面更新后若按钮异常、上传无响应或界面显示异常，请先按 Ctrl+F5 强制刷新；Mac 请按 Command+Shift+R。"
REFRESH_HINT_HTML = "页面更新后若按钮异常、上传无响应或界面显示异常，请先按 <kbd>Ctrl+F5</kbd> 强制刷新；Mac 请按 <kbd>Command+Shift+R</kbd>。"
GLOBAL_NAV_ITEMS = [
    ("/", "fa-home", "首页"),
    ("/dashboard", "fa-gauge-high", "工作台"),
    ("/certificate-translation", "fa-id-card", "证件翻译聚合"),
    ("/pdf2docx", "fa-file-word", "文档预处理"),
    ("/number-check", "fa-check-double", "数字专检"),
    ("/alignment", "fa-object-group", "多语对照"),
    ("/zhongfanyi", "fa-spell-check", "中翻专检"),
]
NAV_ACTIVE_ALIASES = {
    "/certificate-translation": (
        "/certificate-translation",
        "/ocr",
        "/doc-translate",
        "/drivers-license",
        "/business-licence",
    ),
}
APP_SHELL_BOOTSTRAP = f"""
<script>
(() => {{
    try {{
        const params = new URLSearchParams(window.location.search);
        if (params.get('embed') === '1' || window.self !== window.top) {{
            document.documentElement.setAttribute('data-app-embed', '1');
        }}
    }} catch (_) {{}}
}})();
</script>
<style id="appShellStyle">
    html[data-app-embed="1"] .app-shell-slot {{
        display: none !important;
    }}
    html:not([data-app-embed="1"]) header.header .header-nav {{
        display: none !important;
    }}
    html:not([data-app-embed="1"]) header.topbar:not(.unified-global-topbar) {{
        display: none !important;
    }}
    html:not([data-app-embed="1"]) body {{
        padding-top: 0 !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
    }}
    html:not([data-app-embed="1"]) .container {{
        width: min(1280px, calc(100% - 32px)) !important;
    }}
    html:not([data-app-embed="1"]) .app-shell-slot + .page {{
        padding-top: 0 !important;
    }}
    .app-shell-slot {{
        position: relative;
        z-index: 30;
    }}
    .app-shell-inner {{
        width: min(1280px, calc(100% - 32px));
        margin: 24px auto 18px;
    }}
    .unified-global-topbar {{
        display: flex !important;
        justify-content: space-between !important;
        align-items: center !important;
        flex-wrap: wrap !important;
        gap: 16px !important;
        margin-bottom: 18px !important;
        padding-bottom: 18px !important;
        border-bottom: 1px solid rgba(255, 255, 255, 0.06) !important;
    }}
    .unified-global-topbar .brand {{
        display: inline-flex;
        align-items: center;
        gap: 14px;
    }}
    .unified-top-nav {{
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
    }}
    .unified-top-nav a {{
        display: inline-flex;
        align-items: center;
        gap: 8px;
        min-height: 42px;
        padding: 0 16px;
        border-radius: 999px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        background: rgba(255, 255, 255, 0.04);
        text-decoration: none;
        color: #e2e8f0;
        font-size: 14px;
        font-weight: 500;
        transition: all 0.24s ease;
        backdrop-filter: blur(12px);
    }}
    .unified-top-nav a:hover,
    .unified-top-nav a.active {{
        background: rgba(56, 189, 248, 0.16);
        border-color: rgba(56, 189, 248, 0.4);
        color: #fff;
        transform: translateY(-1.5px);
    }}
    .page-hero-header {{
        color: #fff;
        margin-bottom: 24px;
    }}
    .shell-refresh-notice {{
        display: grid;
        grid-template-columns: auto 1fr;
        align-items: flex-start;
        gap: 14px;
        margin-bottom: 20px;
        padding: 16px 18px;
        border-radius: 20px;
        border: 1px solid rgba(251, 191, 36, 0.42);
        background: linear-gradient(135deg, rgba(251, 191, 36, 0.24), rgba(249, 115, 22, 0.2));
        box-shadow: 0 20px 44px rgba(120, 53, 15, 0.2);
        position: relative;
        overflow: hidden;
    }}
    .shell-refresh-notice::before {{
        content: '';
        position: absolute;
        inset: 0 auto 0 0;
        width: 6px;
        background: linear-gradient(180deg, #facc15, #f97316);
    }}
    .shell-refresh-notice .notice-icon {{
        width: 42px;
        height: 42px;
        border-radius: 14px;
        display: grid;
        place-items: center;
        background: rgba(120, 53, 15, 0.18);
        color: #fff7ed;
        box-shadow: inset 0 0 0 1px rgba(255, 247, 237, 0.12);
    }}
    .shell-refresh-notice .notice-copy {{
        min-width: 0;
    }}
    .shell-refresh-notice .notice-badge {{
        display: inline-flex;
        align-items: center;
        min-height: 24px;
        padding: 0 10px;
        margin-bottom: 8px;
        border-radius: 999px;
        background: rgba(120, 53, 15, 0.22);
        color: #fff7ed;
        font-size: 12px;
        font-weight: 800;
        letter-spacing: 0.08em;
    }}
    .shell-refresh-notice .notice-title {{
        display: block;
        margin-bottom: 6px;
        color: #fff7ed;
        font-size: 17px;
        font-weight: 800;
        line-height: 1.35;
    }}
    .shell-refresh-notice .notice-text {{
        color: rgba(255, 247, 237, 0.96);
        font-size: 14px;
        font-weight: 600;
        line-height: 1.7;
    }}
    .shell-refresh-notice kbd {{
        display: inline-flex;
        align-items: center;
        min-height: 26px;
        margin: 0 3px;
        padding: 0 9px;
        border-radius: 8px;
        border: 1px solid rgba(255, 247, 237, 0.32);
        background: rgba(120, 53, 15, 0.3);
        color: #ffffff;
        font-size: 12px;
        font-weight: 800;
        font-family: Consolas, "SFMono-Regular", "Liberation Mono", Menlo, monospace;
        box-shadow: inset 0 -1px 0 rgba(255, 247, 237, 0.16);
        vertical-align: middle;
    }}
    .shell-build-toast {{
        position: fixed;
        right: 18px;
        bottom: 18px;
        z-index: 9999;
        display: flex;
        align-items: center;
        gap: 12px;
        max-width: min(460px, calc(100vw - 28px));
        padding: 14px 16px;
        border-radius: 16px;
        border: 1px solid rgba(125, 211, 252, 0.26);
        background: rgba(15, 23, 42, 0.94);
        box-shadow: 0 24px 50px rgba(2, 8, 23, 0.32);
        color: #f8fafc;
        backdrop-filter: blur(14px);
    }}
    .shell-build-toast button {{
        border: none;
        background: rgba(255, 255, 255, 0.08);
        color: #e2e8f0;
        width: 32px;
        height: 32px;
        border-radius: 999px;
        cursor: pointer;
    }}
    .shell-build-toast button:hover {{
        background: rgba(255, 255, 255, 0.16);
    }}
    select option {{
        background-color: #0d2138;
        color: #f6f8fb;
    }}
    @media (max-width: 720px) {{
        .app-shell-inner {{
            width: min(calc(100% - 20px), 1280px);
            margin: 18px auto 16px;
        }}
        .shell-refresh-notice {{
            border-radius: 16px;
            padding: 14px 14px 14px 16px;
            gap: 12px;
        }}
        .shell-refresh-notice .notice-title {{
            font-size: 15px;
        }}
        .shell-refresh-notice .notice-text {{
            font-size: 13px;
        }}
        .shell-build-toast {{
            right: 10px;
            bottom: 10px;
            left: 10px;
            max-width: none;
        }}
    }}
</style>
"""

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


def _inject_app_shell_bootstrap(html: str) -> str:
    if 'id="appShellStyle"' in html:
        return html
    return html.replace("<head>", f"<head>\n    {APP_SHELL_BOOTSTRAP}", 1)


def _build_app_shell_markup(current_path: str) -> str:
    nav_html = []
    for href, icon, text in GLOBAL_NAV_ITEMS:
        active_prefixes = NAV_ACTIVE_ALIASES.get(href, (href,))
        is_active = any(
            current_path == prefix or (prefix != "/" and current_path.startswith(f"{prefix}/"))
            for prefix in active_prefixes
        )
        active_class = " active" if is_active else ""
        nav_html.append(f'<a href="{href}" class="{active_class.strip()}"><i class="fas {icon}"></i> {text}</a>')

    return f"""
<div id="appShellSlot" class="app-shell-slot">
    <div class="app-shell-inner">
        <header class="topbar unified-global-topbar">
            <div class="brand">
                <div class="brand-mark" style="width: 44px; height: 44px; border-radius: 14px; display: grid; place-items: center; background: linear-gradient(135deg, rgba(56, 189, 248, 0.28), rgba(56, 189, 248, 0.62)); font-size: 20px; color: #fff;">
                    <i class="fas fa-layer-group"></i>
                </div>
                <div style="font-weight: 700; font-size: 17px; color: #fff;">文档处理工作台</div>
            </div>
            <nav class="unified-top-nav">
                {"".join(nav_html)}
            </nav>
        </header>
        <section class="shell-refresh-notice">
            <div class="notice-icon"><i class="fas fa-triangle-exclamation"></i></div>
            <div class="notice-copy">
                <div class="notice-badge">使用提示</div>
                <strong class="notice-title">页面异常时先强制刷新缓存</strong>
                <div class="notice-text">{REFRESH_HINT_HTML}</div>
            </div>
        </section>
    </div>
</div>
"""


def _inject_app_shell_markup(html: str, current_path: str) -> str:
    if 'id="appShellSlot"' in html:
        return html
    return html.replace("<body>", f"<body>\n    {_build_app_shell_markup(current_path)}", 1)


def _render_page(filename: str, current_path: str) -> HTMLResponse:
    content = _read_page(filename)
    content = _inject_app_build_meta(content)
    content = _inject_app_shell_bootstrap(content)
    content = _inject_app_shell_markup(content, current_path)
    content = _inject_static_asset_version(content)
    return HTMLResponse(content=content, headers=HTML_CACHE_HEADERS)


@app.get("/", response_class=HTMLResponse)
async def root():
    return _render_page("nav.html", "/")


@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard_page():
    return _render_page("dashboard.html", "/dashboard")


@app.get("/ocr", response_class=HTMLResponse)
async def ocr_page():
    return _render_page("index.html", "/ocr")


@app.get("/certificate-translation", response_class=HTMLResponse)
async def certificate_translation_page():
    return _render_page("certificate_translation.html", "/certificate-translation")


@app.get("/number-check", response_class=HTMLResponse)
async def number_check_page():
    return _render_page("number_check.html", "/number-check")


@app.get("/alignment", response_class=HTMLResponse)
async def alignment_page():
    return _render_page("alignment.html", "/alignment")


@app.get("/drivers-license", response_class=HTMLResponse)
async def drivers_license_page():
    return _render_page("drivers_license.html", "/drivers-license")


@app.get("/doc-translate", response_class=HTMLResponse)
async def doc_translate_page():
    return _render_page("doc_translate.html", "/doc-translate")


@app.get("/business-licence/embed", response_class=HTMLResponse)
async def business_licence_embed_page():
    return _render_page("business_licence.html", "/business-licence/embed")


@app.get("/business-licence")
async def business_licence_page():
    return RedirectResponse(url="/certificate-translation?tab=business-panel", status_code=307)


@app.get("/zhongfanyi", response_class=HTMLResponse)
async def zhongfanyi_page():
    return _render_page("zhongfanyi.html", "/zhongfanyi")


@app.get("/pdf2docx", response_class=HTMLResponse)
async def pdf2docx_page():
    return _render_page("pdf2docx.html", "/pdf2docx")


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
