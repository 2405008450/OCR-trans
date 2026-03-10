# Docker 部署指南

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
HOST=0.0.0.0
PORT=8001
DEBUG=False
ALLOWED_ORIGINS=*
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

浏览器访问：`http://43.132.156.72:8001`

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
