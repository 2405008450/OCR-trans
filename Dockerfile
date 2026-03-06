# ============================================================
# 基础镜像：Python 3.11 slim
# 云服务器 CPU 版，如有 GPU 可改用 nvidia/cuda 基础镜像
# ============================================================
FROM python:3.11-slim

# 设置工作目录
WORKDIR /app

# 环境变量
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PADDLEOCR_HOME=/app/.paddleocr \
    OPENCV_IO_MAX_IMAGE_PIXELS=1099511627776 \
    # 让 pip 跳过 SSL 验证（代理环境下避免 SSL EOF 错误）
    PIP_TRUSTED_HOST="pypi.org files.pythonhosted.org download.pytorch.org pypi.tuna.tsinghua.edu.cn www.paddlepaddle.org.cn"

# 安装系统依赖（绕过代理直连 apt 源）
RUN unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY no_proxy NO_PROXY; \
    apt-get update && apt-get install -y --no-install-recommends \
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

# ── 第一层：升级 pip，写入全局 pip 配置（代理 TUN 模式下禁用 SSL 验证）──
RUN pip install --upgrade pip && \
    mkdir -p /root/.config/pip && \
    printf '[global]\ntrusted-host = pypi.org\n    files.pythonhosted.org\n    download.pytorch.org\n    www.paddlepaddle.org.cn\n' \
    > /root/.config/pip/pip.conf

# ── 第二层：安装 PyTorch（CPU 版，体积大，单独缓存）──
RUN pip install torch==2.10.0 torchvision==0.25.0 \
        --index-url https://download.pytorch.org/whl/cpu

# ── 第三层：安装 PaddlePaddle（CPU 版）──
RUN pip install paddlepaddle==3.2.2 \
    -i https://www.paddlepaddle.org.cn/packages/stable/cpu/

# ── 第四层：安装应用依赖 ──
COPY requirements.txt .
# zai-sdk 的 pyjwt 依赖声明与 zhipuai 冲突，单独用 --no-deps 安装（实际运行兼容 2.8.0）
RUN pip install -r requirements.txt \
    && pip install zai-sdk==0.2.2 --no-deps

# ── 复制项目代码 ──
COPY app/ ./app/
COPY static/ ./static/
COPY businesslicence/ ./businesslicence/
COPY memory/ ./memory/
COPY 专检/ ./专检/

# 创建运行时目录
RUN mkdir -p uploads outputs temp_images \
    businesslicence/uploads businesslicence/outputs \
    memory/Result_Output

# 暴露端口
EXPOSE 8001

# 启动命令
CMD ["python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8001"]
