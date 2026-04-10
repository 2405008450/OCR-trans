# Docker 部署指南
uvicorn app.main:app --host 127.0.0.1 --port 8001 --proxy-headers --forwarded-allow-ips='*'
目标服务器 IP：`43.160.215.225`，端口：`8001`

---

## 一、本地准备（Windows 开发机）

### 1. 安装 Docker Desktop
如果本地还没安装：https://www.docker.com/products/docker-desktop/

### 2. 导出 conda 环境的完整依赖（可选，用于核对）
```bash
conda run -n ocr_lama pip freeze > requirements_full.txt
```

### 3. 在本地构建镜像（可选，网络好时也可直接在服务器构建）
```bash
cd e:/fastapi-llm-demo
docker build -t fastapi-llm-demo:latest .
```
> 首次构建约需 10-30 分钟（需下载基础镜像 ~5GB）

### 4. 打包镜像为文件上传到服务器（如果网络慢）
```bash
# 导出镜像
docker save fastapi-llm-demo:latest | gzip > fastapi-llm-demo.tar.gz
# 上传到服务器（约 3-6 GB）
scp fastapi-llm-demo.tar.gz root@43.160.215.225:/root/
```

---

## 二、上传项目代码到服务器

```bash
# 方法1：使用 scp 上传整个项目（排除大文件）
scp -r e:/fastapi-llm-demo root@43.160.215.225:/root/fastapi-llm-demo

# 方法2：用 git（推荐，如果服务器能访问 git 仓库）
# 在服务器上执行：
# git clone <你的仓库地址> /root/fastapi-llm-demo
```

---

## 三、在云服务器部署

### 1. SSH 登录服务器
```bash
ssh root@43.160.215.225
```

### 2. 安装 Docker（如果服务器没有）
```bash
# Ubuntu/Debian
curl -fsSL https://get.docker.com | bash
systemctl enable docker
systemctl start docker

# 安装 docker-compose
sudo apt-get install -y docker-compose-plugin
# 或者
sudo curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-linux-x86_64" -o /usr/local/bin/docker-compose
sudo chmod +x /usr/local/bin/docker-compose
```

### 3. 进入项目目录
```bash
cd /root/fastapi-llm-demo
```

### 4. 如果是上传了镜像文件，先导入
```bash
docker load < fastapi-llm-demo.tar.gz
```

### 5. 配置环境变量
```bash
# 复制并修改 .env 文件
cp env.example .env
nano .env
# 确保填写正确的 DEEPSEEK_API_KEY
```

`.env` 文件内容：
```
DEEPSEEK_API_KEY=你的真实API_KEY
DEEPSEEK_BASE_URL=https://api.deepseek.com
GOOGLE_API_KEY=????GOOGLE_API_KEY
GEMINI_DEFAULT_ROUTE=google
OPENROUTER_API_KEY=你的真实OPENROUTER_API_KEY
OPENROUTER_BASE_URL=https://openrouter.ai/api/v1
LIBREOFFICE_PATH=/usr/bin/soffice
HOST=0.0.0.0
PORT=8001
DEBUG=False
ALLOWED_ORIGINS=*
```

### 5.1 LibreOffice 说明
当前项目的 PDF/图片转 Word 功能依赖 LibreOffice 的 `soffice` 命令。

- Docker 部署：`Dockerfile` 已安装 `libreoffice` 和 `libreoffice-writer`，默认路径为 `/usr/bin/soffice`
- 裸机 Linux 部署：请先安装 LibreOffice，再执行 `export LIBREOFFICE_PATH=/usr/bin/soffice`
- 验证命令：
```bash
which soffice
soffice --headless --version
```

### 6. 如果服务器直接构建（网络好的情况）
```bash
cd /root/fastapi-llm-demo
docker build -t fastapi-llm-demo:latest .
```

### 7. 启动服务
```bash
docker-compose up -d
```

### 8. 查看运行状态
```bash
docker-compose logs -f        # 实时查看日志
docker-compose ps             # 查看容器状态
```

---

## 四、验证部署

浏览器访问：`http://43.160.215.225:8001`

---

## 五、常用运维命令

```bash
# 停止服务
docker-compose down

# 重启服务
docker-compose restart

# 更新代码后重新构建并重启
docker-compose down
docker build -t fastapi-llm-demo:latest .
docker-compose up -d

# 进入容器调试
docker exec -it fastapi-llm bash

# 查看容器内日志
docker logs fastapi-llm --tail 100
```

---

## 六、防火墙配置

确保云服务器安全组放开 **8001 端口（TCP）**：
- 登录云服务器控制台（腾讯云/阿里云）
- 找到"安全组" → 入站规则 → 添加规则
- 协议：TCP，端口：8001，来源：0.0.0.0/0

---

## 七、GPU vs CPU 说明

| 情况 | Dockerfile 基础镜像 | 性能 |
|------|-------------------|------|
| 服务器有 NVIDIA GPU | `paddlepaddle/paddle:3.0.0-gpu-cuda11.8-cudnn8.6-trt8.5` | 快 |
| 服务器无 GPU（纯CPU） | `paddlepaddle/paddle:3.0.0` | 慢但能用 |

如果服务器是 **CPU 型**，编辑 `Dockerfile` 第4行改为：
```dockerfile
FROM paddlepaddle/paddle:3.0.0
```
同时 `requirements_docker.txt` 中不需要 nvidia 相关包。

---

## 八、镜像大小优化说明

完整镜像约 **6-8 GB**（含 CUDA、PaddlePaddle、torch 等），这是正常的。
若要减小体积，可以：
1. 使用 CPU 版基础镜像（约 3-4 GB）
2. 去掉 `simple-lama-inpainting`（会去掉 torch 依赖）

---

## 九、多进程部署与性能优化（防卡死指南）

如果你发现**同时运行两个不同项目/任务时，系统直接卡死，Docker 日志停止跳动**，这通常不是单纯的 FastAPI 因为并发阻塞导致的，更核心的原因在于 **服务器资源耗尽（尤其是内存 OOM）**。

### 1. 为什么会卡死？
当前项目使用了深度学习模型（如 `PaddleOCR`、`llama`、`torch`）。这些 AI 模型不仅吃 CPU 计算力，更占用大量**内存 (RAM)**。
在单进程模式下，当有请求进来，主进程满负荷运行 AI 推理，此时如果有新的请求进来，FastAPI 事件循环会被阻塞，表现为"卡住排队等待"。
一旦开启多进程，每个进程都会独立在内存中加载一份完整的 AI 模型体系（Paddle、Torch），如果你的服务器是 2核心且内存较小（例如 4核 8G 或者 2核 4G），内存会被迅速撑爆！此时 Linux 系统的 OOM (Out of Memory) Killer 会直接杀死你的进程，或者疯狂触发 Swap 磁盘交换，导致整个云服务器陷入硬盘 I/O 死锁，表现也就是 **完全卡死不动**。

对于 2核心 的服务器，AI 推理的算力已经属于极度紧缺状态，**强烈建议不建议盲目增加部署进程数**。

### 2. 如何配置文件开启多进程？

如果你已经确认你的服务器**内存绝对足够充足（例如 16GB 以上）**，并且非常希望通过多进程解决排队阻塞，可以修改项目的 `Dockerfile` 启动命令。

编辑 `Dockerfile` 最新的一行：
```dockerfile
# 将原本的单进程启动：
CMD ["python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8001"]

# 修改为（增加 --workers 参数，例如开启 2 个进程）：
CMD ["python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8001", "--workers", "2"]
```
如果你使用了 `gunicorn`，也可以修改为：
```bash
CMD ["gunicorn", "app.main:app", "-w", "2", "-k", "uvicorn.workers.UvicornWorker", "-b", "0.0.0.0:8001", "--timeout", "300"]
```

### 3. 给 2核服务器（或低配服务器）的终极建议
由于你是 2核 服务器，**强开多进程必然导致物理崩溃**。建议采取以下方案应对多项目卡死：
1. **继续使用单进程** `uvicorn`，但是接受排队。可以通过增加前端的“排队提示”来让用户知道系统正在处理上一个任务。
2. **异步包裹同步计算**: 确保代码中执行 `PaddleOCR` 或 `Llama` 推理的方法使用 `asyncio.to_thread` 或 `fastapi.concurrency.run_in_threadpool` 放到后台线程池执行，这样不仅主线程可以继续受理其他请求（不会卡死），而且也不会重复消耗过多内存。
3. **扩大 Swap 分区**: 在云服务器上配置 8GB 的虚拟内存（Swap），虽然推理会变慢，但可以避免内存不够导致的直接假死。
