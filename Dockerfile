FROM python:3.11-slim

WORKDIR /app

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PADDLEOCR_HOME=/app/.paddleocr \
    LIBREOFFICE_PATH=/usr/bin/soffice \
    OPENCV_IO_MAX_IMAGE_PIXELS=1099511627776 \
    PIP_TRUSTED_HOST="pypi.org files.pythonhosted.org download.pytorch.org pypi.tuna.tsinghua.edu.cn www.paddlepaddle.org.cn"

RUN unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY no_proxy NO_PROXY; \
    apt-get update && apt-get install -y --no-install-recommends \
    libglib2.0-0 \
    libgl1 \
    libgomp1 \
    libreoffice \
    libreoffice-writer \
    fonts-noto-cjk \
    libsm6 \
    libxext6 \
    libxrender-dev \
    libgstreamer1.0-0 \
    python3-tk \
    tk-dev \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --upgrade pip && \
    mkdir -p /root/.config/pip && \
    printf '[global]\ntrusted-host = pypi.org\n    files.pythonhosted.org\n    download.pytorch.org\n    www.paddlepaddle.org.cn\n' \
    > /root/.config/pip/pip.conf

RUN pip install torch==2.10.0 torchvision==0.25.0 \
        --index-url https://download.pytorch.org/whl/cpu

RUN pip install paddlepaddle==3.2.2 \
    -i https://www.paddlepaddle.org.cn/packages/stable/cpu/

COPY requirements.txt .
RUN pip install -r requirements.txt \
    && pip install PyJWT==2.11.0 --no-deps \
    && pip install zai-sdk==0.2.2 --no-deps

COPY app/ ./app/
COPY pdf2docx.py ./
COPY static/ ./static/
COPY businesslicence/ ./businesslicence/
COPY Driver's_License/ ./Driver's_License/
COPY memory/ ./memory/
COPY 专检/ ./专检/

RUN mkdir -p uploads outputs temp_images \
    businesslicence/uploads businesslicence/outputs \
    memory/Result_Output

EXPOSE 8001

CMD ["python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8001"]


