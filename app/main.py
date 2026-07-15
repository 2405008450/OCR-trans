from pathlib import Path

import re
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from app.controller import task
from app.core.config import settings
from app.core.request_context import reset_client_ip, set_client_ip
from app.db.init_db import init_db
from app.service.task_queue_service import task_queue_service

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
RELEASE_NOTES = [
    "7月15日_上线合并 PDF，支持共享目录扫描、排序选择与本地合并下载",
    "6月16日_数字专检新增 PDF 支持",
    "6月30日_新增重复任务校验，批量导出任务支持持久化保留",
    "7月6日_文档预处理优化聊天软件截图与复杂图片 OCR",
    "7月8日_上线字数统计，支持共享路径批量扫描并导出 Excel / JSON",
    "7月14日_上线 MSG 转 Word / PDF，字数统计新增 CAD 与 LLM OCR 支持",
]
PRIMARY_NAV_ITEMS = [
    ("/", "fa-home", "首页"),
    ("/dashboard", "fa-gauge-high", "工作台"),
]
TOOL_NAV_GROUPS = [
    ("翻译与识别", "fa-language", [
        ("/certificate-translation", "fa-id-card", "证件翻译", "驾驶证、营业执照及通用证件"),
        ("/alignment", "fa-object-group", "多语对照", "原译文对齐与语料沉淀"),
    ]),
    ("转换与处理", "fa-shuffle", [
        ("/pdf2docx", "fa-file-word", "文档预处理", "PDF / 图片转可编辑 Word"),
        ("/msg-convert", "fa-envelope-open-text", "MSG 转文档", "邮件转 Word / PDF，保留正文与内嵌图片"),
        ("/pdf-merge", "fa-object-group", "合并 PDF", "共享目录选取、排序并合并 PDF"),
        ("/word-count", "fa-calculator", "字数统计", "批量扫描并生成统计报告"),
    ]),
    ("检查与校对", "fa-shield-halved", [
        ("/number-check", "fa-check-double", "数字专检", "双语文档数字一致性检查"),
        ("/zhongfanyi", "fa-spell-check", "中翻专检", "规则与 AI 联合审校"),
    ]),
]
NAV_ACTIVE_ALIASES = {
    "/certificate-translation": (
        "/certificate-translation",
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
        const isEmbedded = params.get('embed') === '1' || window.self !== window.top;
        if (isEmbedded) {{
            document.documentElement.setAttribute('data-app-embed', '1');
        }} else {{
            document.documentElement.classList.add('app-transition-enabled', 'app-page-preparing');
            window.setTimeout(() => {{
                document.documentElement.classList.remove('app-page-preparing');
            }}, 2500);
        }}
    }} catch (_) {{}}
}})();
</script>
<style id="appShellStyle">
    html.app-transition-enabled {{
        background: #040812;
    }}
    html.app-page-preparing body {{
        opacity: 0;
    }}
    html.app-page-ready body {{
        opacity: 1;
    }}
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
        flex-wrap: nowrap !important;
        gap: 16px !important;
        margin-bottom: 18px !important;
        padding-bottom: 18px !important;
        border-bottom: 1px solid rgba(255, 255, 255, 0.06) !important;
        position: relative;
    }}
    .unified-global-topbar .brand {{
        display: inline-flex;
        align-items: center;
        gap: 14px;
    }}
    .unified-top-nav {{
        display: flex;
        align-items: center;
        gap: 10px;
    }}
    .unified-top-nav a {{
        display: inline-flex;
        align-items: center;
        gap: 8px;
        min-height: 42px;
        line-height: 1;
        padding: 0 16px;
        border-radius: 999px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        background: rgba(255, 255, 255, 0.04);
        text-decoration: none;
        color: #e2e8f0;
        font-size: 14px;
        font-weight: 500;
        transition: background-color 0.18s ease, border-color 0.18s ease, color 0.18s ease;
        backdrop-filter: blur(12px);
    }}
    .unified-top-nav a:hover,
    .unified-top-nav a.active {{
        background: rgba(56, 189, 248, 0.16);
        border-color: rgba(56, 189, 248, 0.4);
        color: #fff;
    }}
    .tool-menu {{
        position: relative;
    }}
    .tool-menu summary {{
        display: inline-flex;
        align-items: center;
        gap: 9px;
        min-height: 42px;
        padding: 0 16px;
        border-radius: 999px;
        border: 1px solid rgba(56, 189, 248, 0.28);
        background: rgba(56, 189, 248, 0.1);
        color: #e0f2fe;
        cursor: pointer;
        font-size: 14px;
        font-weight: 700;
        list-style: none;
        user-select: none;
        white-space: nowrap;
    }}
    .tool-menu summary::-webkit-details-marker {{
        display: none;
    }}
    .tool-menu summary::after {{
        content: '\\f078';
        font-family: "Font Awesome 6 Free";
        font-size: 11px;
        font-weight: 900;
        transition: transform 0.18s ease;
    }}
    .tool-menu[open] summary,
    .tool-menu.is-active summary {{
        border-color: rgba(56, 189, 248, 0.58);
        background: rgba(56, 189, 248, 0.18);
        color: #fff;
    }}
    .tool-menu[open] summary::after {{
        transform: rotate(180deg);
    }}
    .tool-menu-panel {{
        position: absolute;
        top: calc(100% + 12px);
        right: 0;
        z-index: 80;
        width: min(760px, calc(100vw - 32px));
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 10px;
        padding: 14px;
        border: 1px solid rgba(125, 211, 252, 0.2);
        border-radius: 22px;
        background: rgba(7, 19, 34, 0.98);
        box-shadow: 0 28px 80px rgba(2, 8, 23, 0.56);
        backdrop-filter: blur(22px);
    }}
    .tool-menu-group {{
        min-width: 0;
        padding: 8px;
        border-radius: 16px;
        background: rgba(255, 255, 255, 0.025);
    }}
    .tool-menu-title {{
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 4px 6px 10px;
        color: #7dd3fc;
        font-size: 12px;
        font-weight: 800;
        letter-spacing: 0.06em;
    }}
    .unified-top-nav .tool-menu-link {{
        display: grid;
        grid-template-columns: 34px minmax(0, 1fr);
        min-height: 64px;
        padding: 10px;
        border: 0;
        border-radius: 13px;
        background: transparent;
    }}
    .tool-menu-link > i {{
        width: 34px;
        height: 34px;
        display: grid;
        place-items: center;
        border-radius: 10px;
        background: rgba(56, 189, 248, 0.1);
        color: #7dd3fc;
    }}
    .tool-menu-link strong,
    .tool-menu-link small {{
        display: block;
    }}
    .tool-menu-link strong {{
        margin-bottom: 4px;
        color: #f8fafc;
        font-size: 14px;
    }}
    .tool-menu-link small {{
        overflow: hidden;
        color: rgba(226, 232, 240, 0.62);
        font-size: 11px;
        line-height: 1.4;
        text-overflow: ellipsis;
        white-space: nowrap;
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
    .shell-release-note {{
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 14px;
        margin-bottom: 18px;
        padding: 14px 18px;
        border-radius: 14px;
        border: 1px solid rgba(14, 165, 233, 0.18);
        background: linear-gradient(135deg, rgba(11, 31, 52, 0.6), rgba(6, 17, 29, 0.5));
        box-shadow: 0 12px 28px rgba(0, 0, 0, 0.2);
        position: relative;
        overflow: hidden;
    }}
    .shell-release-note::before {{
        content: '';
        position: absolute;
        inset: 0 0 auto 0;
        height: 4px;
        background: linear-gradient(90deg, #0ea5e9, #38bdf8);
    }}
    .shell-release-note .release-icon {{
        width: 40px;
        height: 40px;
        border-radius: 12px;
        display: grid;
        place-items: center;
        background: linear-gradient(135deg, #0284c7, #0369a1);
        color: #ffffff;
        font-size: 18px;
        flex: 0 0 auto;
    }}
    .shell-release-note .release-copy {{
        min-width: 0;
        flex: 1 1 0;
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 10px 12px;
    }}
    .shell-release-note .release-badge {{
        display: inline-flex;
        align-items: center;
        min-height: 24px;
        padding: 0 10px;
        margin: 0;
        border-radius: 999px;
        background: rgba(14, 165, 233, 0.18) !important;
        border: 1px solid rgba(14, 165, 233, 0.32) !important;
        color: #e0f2fe !important;
        -webkit-text-fill-color: currentColor !important;
        font-size: 12px;
        font-weight: 900;
        letter-spacing: 0.06em;
        opacity: 1 !important;
        flex: 0 0 auto;
    }}
    .shell-release-note .release-title {{
        display: inline;
        margin: 0;
        color: #f0f9ff !important;
        -webkit-text-fill-color: currentColor !important;
        font-size: 15px;
        font-weight: 800;
        line-height: 1.5;
        text-shadow: none;
        opacity: 1 !important;
        flex: 0 1 auto;
    }}
    .shell-release-note .release-list {{
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 8px;
        margin: 0;
        padding: 0;
        list-style: none;
        flex: 1 1 100%;
    }}
    .shell-release-note .release-list li {{
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 6px 10px;
        border-radius: 999px;
        background: rgba(56, 189, 248, 0.12) !important;
        border: 1px solid rgba(56, 189, 248, 0.16) !important;
        color: #bae6fd !important;
        -webkit-text-fill-color: currentColor !important;
        opacity: 1 !important;
        font-size: 14px;
        font-weight: 700;
        line-height: 1.4;
        text-shadow: none;
    }}
    .shell-release-note .release-list li::before {{
        content: '';
        width: 6px;
        height: 6px;
        border-radius: 999px;
        background: #38bdf8;
        flex: 0 0 auto;
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
        .unified-global-topbar {{
            align-items: flex-start !important;
            flex-wrap: wrap !important;
        }}
        .unified-global-topbar .brand {{
            flex: 1 1 100%;
        }}
        .unified-top-nav {{
            display: grid;
            grid-template-columns: 1fr 1fr 1.35fr;
            width: 100%;
        }}
        .unified-top-nav > a {{
            justify-content: center;
            min-width: 0;
            padding: 0 10px;
        }}
        .tool-menu {{
            width: 100%;
        }}
        .tool-menu summary {{
            justify-content: center;
            width: 100%;
            padding: 0 10px;
        }}
        .tool-menu-panel {{
            position: fixed;
            top: 132px;
            right: 10px;
            left: 10px;
            width: auto;
            grid-template-columns: 1fr;
            max-height: min(66vh, 520px);
            overflow-y: auto;
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
        .shell-release-note {{
            align-items: flex-start;
            border-radius: 12px;
            padding: 14px 14px 14px 16px;
            gap: 12px;
        }}
        .shell-release-note .release-copy {{
            display: grid;
            gap: 8px;
        }}
        .shell-release-note .release-title {{
            font-size: 14px;
        }}
        .shell-release-note .release-list {{
            gap: 7px;
        }}
        .shell-release-note .release-list li {{
            font-size: 13px;
            padding: 6px 9px;
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


@app.middleware("http")
async def capture_client_ip(request: Request, call_next):
    token = set_client_ip(request.client.host if request.client else None)
    try:
        return await call_next(request)
    finally:
        reset_client_ip(token)

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
    await task_queue_service.start()


@app.on_event("shutdown")
async def shutdown_event():
    await task_queue_service.stop()


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
    primary_nav_html = []
    for href, icon, text in PRIMARY_NAV_ITEMS:
        active_prefixes = NAV_ACTIVE_ALIASES.get(href, (href,))
        is_active = any(
            current_path == prefix or (prefix != "/" and current_path.startswith(f"{prefix}/"))
            for prefix in active_prefixes
        )
        active_class = " active" if is_active else ""
        primary_nav_html.append(f'<a href="{href}" class="{active_class.strip()}"><i class="fas {icon}"></i> {text}</a>')

    tool_groups_html = []
    tool_menu_active = False
    for group_name, group_icon, items in TOOL_NAV_GROUPS:
        item_html = []
        for href, icon, text, description in items:
            active_prefixes = NAV_ACTIVE_ALIASES.get(href, (href,))
            is_active = any(
                current_path == prefix or (prefix != "/" and current_path.startswith(f"{prefix}/"))
                for prefix in active_prefixes
            )
            tool_menu_active = tool_menu_active or is_active
            active_class = " active" if is_active else ""
            item_html.append(
                f'<a href="{href}" class="tool-menu-link{active_class}">'
                f'<i class="fas {icon}"></i><span><strong>{text}</strong><small>{description}</small></span></a>'
            )
        tool_groups_html.append(
            f'<section class="tool-menu-group"><div class="tool-menu-title">'
            f'<i class="fas {group_icon}"></i>{group_name}</div>{"".join(item_html)}</section>'
        )

    release_note_items = "".join(f"<li>{item}</li>" for item in RELEASE_NOTES)

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
                {"".join(primary_nav_html)}
                <details class="tool-menu{' is-active' if tool_menu_active else ''}">
                    <summary><i class="fas fa-border-all"></i> 工具中心 <span>7</span></summary>
                    <div class="tool-menu-panel">{"".join(tool_groups_html)}</div>
                </details>
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
        <section class="shell-release-note">
            <div class="release-icon"><i class="fas fa-bullhorn"></i></div>
            <div class="release-copy">
                <div class="release-badge">最近更新</div>
                <strong class="release-title">当前版本已同步以下内容</strong>
                <ul class="release-list">
                    {release_note_items}
                </ul>
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


@app.get("/msg-convert", response_class=HTMLResponse)
async def msg_convert_page():
    return _render_page("msg_convert.html", "/msg-convert")


@app.get("/word-count", response_class=HTMLResponse)
async def word_count_page():
    return _render_page("word_count.html", "/word-count")


@app.get("/pdf-merge", response_class=HTMLResponse)
async def pdf_merge_page():
    return _render_page("pdf_merge.html", "/pdf-merge")


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
