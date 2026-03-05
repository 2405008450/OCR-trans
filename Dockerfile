# ============================================================
# 基础镜像：PaddlePaddle GPU 版（含 CUDA 11.8 + cuDNN 8）
# 如果云服务器没有GPU，改用：paddlepaddle/paddle:3.0.0 (CPU版)
# ============================================================
FROM python:3.10-slim

# 设置工作目录
WORKDIR /app

# 设置环境变量
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    # PaddleOCR 模型缓存目录
    PADDLEOCR_HOME=/app/.paddleocr \
    # 避免 OpenCV 的 GUI 依赖问题
    OPENCV_IO_MAX_IMAGE_PIXELS=1099511627776

# 安装系统依赖（OpenCV、字体支持等）
RUN apt-get update && apt-get install -y --no-install-recommends \
    libglib2.0-0 \
    libgl1 \
    libgomp1 \
    fonts-noto-cjk \
    libsm6 \
    libxext6 \
    libxrender-dev \
    libgstreamer1.0-0 \
    python3-tk \
    tk-dev \
    && rm -rf /var/lib/apt/lists/*

# 先复制依赖文件（利用 Docker 层缓存）
COPY requirements_docker.txt .

# 安装 Python 依赖（优先用清华源加速）
RUN pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple && \
    pip install -r requirements_docker.txt \
        -i https://pypi.tuna.tsinghua.edu.cn/simple \
        --extra-index-url https://pypi.org/simple

# 复制项目代码
COPY app/ ./app/
COPY static/ ./static/
COPY businesslicence/ ./businesslicence/
# memory.py 被 alignment_service.py 动态加载，必须复制
COPY memory/ ./memory/
# 专检目录含 llm 子模块，需一并复制
COPY 专检/ ./专检/

# 创建必要的目录
RUN mkdir -p uploads outputs temp_images \
    businesslicence/uploads businesslicence/outputs \
    memory/Result_Output

# 复制环境变量模板（实际运行时通过 -e 或 .env 文件注入）
COPY env.example .env

# 暴露端口
EXPOSE 8001

# 启动命令
CMD ["python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8001"]
