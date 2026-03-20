from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
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


@app.get("/", response_class=HTMLResponse)
async def root():
    return _read_page("nav.html")


@app.get("/ocr", response_class=HTMLResponse)
async def ocr_page():
    return _read_page("index.html")


@app.get("/number-check", response_class=HTMLResponse)
async def number_check_page():
    return _read_page("number_check.html")


@app.get("/alignment", response_class=HTMLResponse)
async def alignment_page():
    return _read_page("alignment.html")


@app.get("/doc-translate", response_class=HTMLResponse)
async def doc_translate_page():
    return _read_page("doc_translate.html")


@app.get("/zhongfanyi", response_class=HTMLResponse)
async def zhongfanyi_page():
    return _read_page("zhongfanyi.html")


@app.get("/pdf2docx", response_class=HTMLResponse)
async def pdf2docx_page():
    return _read_page("pdf2docx.html")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "app.main:app",
        host=settings.HOST,
        port=settings.PORT,
        reload=settings.DEBUG_ENABLED,
    )
