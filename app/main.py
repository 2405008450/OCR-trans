from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from app.controller import task
from app.core.config import settings
import os
from pathlib import Path

# 获取项目根目录（确保路径正确）
BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "static"
UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUTS_DIR = BASE_DIR / "outputs"
TEMP_IMAGES_DIR = BASE_DIR / "temp_images"

app = FastAPI(
    title="图片OCR翻译系统",
    description="AI驱动的智能文档识别与翻译",
    version="1.0.0"
)

# 配置CORS（允许跨域访问，云服务器部署时可能需要）
allowed_origins = settings.ALLOWED_ORIGINS.split(",") if settings.ALLOWED_ORIGINS != "*" else ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 挂载静态文件目录（使用绝对路径）
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
    print(f"✅ 静态文件目录已挂载: {STATIC_DIR}")
else:
    print(f"⚠️  警告：静态文件目录不存在: {STATIC_DIR}")

# 挂载输出文件目录（允许前端访问处理结果，使用绝对路径）
if UPLOADS_DIR.exists():
    app.mount("/uploads", StaticFiles(directory=str(UPLOADS_DIR)), name="uploads")
    print(f"✅ 上传目录已挂载: {UPLOADS_DIR}")
if OUTPUTS_DIR.exists():
    app.mount("/outputs", StaticFiles(directory=str(OUTPUTS_DIR)), name="outputs")
    print(f"✅ 输出目录已挂载: {OUTPUTS_DIR}")
if TEMP_IMAGES_DIR.exists():
    app.mount("/temp_images", StaticFiles(directory=str(TEMP_IMAGES_DIR)), name="temp_images")
    print(f"✅ 临时图片目录已挂载: {TEMP_IMAGES_DIR}")

# 注册任务路由
app.include_router(task.router)

# 根路径返回前端页面
@app.get("/", response_class=HTMLResponse)
async def root():
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        with open(index_file, "r", encoding="utf-8") as f:
            return f.read()
    else:
        return f"<h1>错误：找不到 index.html</h1><p>路径: {index_file}</p>"

# 启动（一条命令，局域网访问 http://192.168.31.125:8001）：
#   python -m app.main
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "app.main:app",
        host=settings.HOST,
        port=settings.PORT,
        reload=settings.DEBUG,
    )
