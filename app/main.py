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
BL_OUTPUTS_DIR = BASE_DIR / "businesslicence" / "outputs"

app = FastAPI(
    title="图片OCR翻译系统",
    description="AI驱动的智能文档识别与翻译",
    version="1.0.0"
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

BL_OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
app.mount("/bl-outputs", StaticFiles(directory=str(BL_OUTPUTS_DIR)), name="bl_outputs")

app.include_router(task.router)


@app.on_event("startup")
async def startup_event():
    init_db()
    await ocr_task_queue.start()


@app.on_event("shutdown")
async def shutdown_event():
    await ocr_task_queue.stop()


@app.get("/", response_class=HTMLResponse)
async def root():
    nav_file = STATIC_DIR / "nav.html"
    if nav_file.exists():
        with open(nav_file, "r", encoding="utf-8") as f:
            return f.read()
    return f"<h1>错误：找不到 nav.html</h1><p>路径: {nav_file}</p>"


@app.get("/ocr", response_class=HTMLResponse)
async def ocr_page():
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        with open(index_file, "r", encoding="utf-8") as f:
            return f.read()
    return f"<h1>错误：找不到 index.html</h1><p>路径: {index_file}</p>"


@app.get("/number-check", response_class=HTMLResponse)
async def number_check_page():
    page_file = STATIC_DIR / "number_check.html"
    if page_file.exists():
        with open(page_file, "r", encoding="utf-8") as f:
            return f.read()
    return f"<h1>错误：找不到 number_check.html</h1><p>路径: {page_file}</p>"


@app.get("/alignment", response_class=HTMLResponse)
async def alignment_page():
    page_file = STATIC_DIR / "alignment.html"
    if page_file.exists():
        with open(page_file, "r", encoding="utf-8") as f:
            return f.read()
    return f"<h1>错误：找不到 alignment.html</h1><p>路径: {page_file}</p>"


@app.get("/business-licence", response_class=HTMLResponse)
async def business_licence_page():
    page_file = STATIC_DIR / "business_licence.html"
    if page_file.exists():
        with open(page_file, "r", encoding="utf-8") as f:
            return f.read()
    return f"<h1>错误：找不到 business_licence.html</h1><p>路径: {page_file}</p>"


@app.get("/zhongfanyi", response_class=HTMLResponse)
async def zhongfanyi_page():
    page_file = STATIC_DIR / "zhongfanyi.html"
    if page_file.exists():
        with open(page_file, "r", encoding="utf-8") as f:
            return f.read()
    return f"<h1>错误：找不到 zhongfanyi.html</h1><p>路径: {page_file}</p>"


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "app.main:app",
        host=settings.HOST,
        port=settings.PORT,
        reload=settings.DEBUG,
    )

