# 腾讯云服务器部署指南

## 📋 部署前准备

### 1. 服务器要求

- **操作系统**: Ubuntu 20.04+ / CentOS 7+ / Debian 10+
- **Python版本**: Python 3.8+
- **内存**: 建议 4GB+（OCR处理需要较多内存）
- **磁盘**: 建议 20GB+（模型文件较大）
- **网络**: 需要公网IP

### 2. 本地准备

#### 打包项目文件

```bash
# 创建部署包（排除不必要的文件）
tar -czf fastapi-ocr-deploy.tar.gz \
    --exclude='.venv' \
    --exclude='__pycache__' \
    --exclude='*.pyc' \
    --exclude='.git' \
    --exclude='.idea' \
    --exclude='data/*.db' \
    --exclude='uploads/*' \
    --exclude='outputs/*' \
    --exclude='temp_images/*' \
    app/ static/ requirements.txt .env.example \
    start_cloud.sh CLOUD_DEPLOYMENT.md
```

#### Windows 打包（PowerShell）

```powershell
# 使用 7-Zip 或 WinRAR 打包
# 排除以下目录/文件：
# - .venv
# - __pycache__
# - .git
# - .idea
# - data/*.db
# - uploads/*
# - outputs/*
# - temp_images/*
```

## 🚀 部署步骤

### 步骤 1: 上传文件到服务器

#### 方法 A: 使用 SCP

```bash
scp fastapi-ocr-deploy.tar.gz root@your_server_ip:/root/
```

#### 方法 B: 使用 SFTP 工具

- WinSCP (Windows)
- FileZilla (跨平台)
- VS Code Remote (推荐)

### 步骤 2: 登录服务器

```bash
ssh root@your_server_ip
```

### 步骤 3: 解压并安装依赖

```bash
# 创建项目目录
mkdir -p /opt/fastapi-ocr
cd /opt/fastapi-ocr

# 解压文件
tar -xzf ~/fastapi-ocr-deploy.tar.gz

# 安装系统依赖（Ubuntu/Debian需要）
sudo apt-get update
sudo apt-get install -y \
    python3-venv \
    python3-pip \
    libgl1-mesa-glx \
    libglib2.0-0 \
    libsm6 \
    libxext6 \
    libxrender-dev \
    libgomp1

# 创建虚拟环境
python3 -m venv .venv
source .venv/bin/activate

# 升级pip
pip install --upgrade pip

# 安装依赖（可能需要较长时间，PaddleOCR模型较大）
# 如果网络慢，可以使用国内镜像：
# pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install -r requirements.txt
```

### 步骤 4: 配置环境变量

```bash
# 复制示例文件
cp .env.example .env

# 编辑配置文件
nano .env
```

**重要配置项：**

```env
# DeepSeek API密钥（必须修改）
DEEPSEEK_API_KEY=your_actual_api_key

# 服务器配置
HOST=0.0.0.0
PORT=8000
DEBUG=False

# 云服务器公网IP（用于显示访问地址）
SERVER_IP=your_server_public_ip

# CORS配置（生产环境建议限制）
ALLOWED_ORIGINS=https://yourdomain.com,https://www.yourdomain.com
```

### 步骤 5: 初始化数据库

```bash
source .venv/bin/activate
python -m app.db.init_db
```

### 步骤 6: 配置防火墙

#### Ubuntu/Debian (ufw)

```bash
# 安装ufw（如果未安装）
apt-get update
apt-get install -y ufw

# 开放8000端口
ufw allow 8000/tcp
ufw enable
ufw status
```

#### CentOS/RHEL (firewalld)

```bash
# 开放8000端口
firewall-cmd --permanent --add-port=8000/tcp
firewall-cmd --reload
firewall-cmd --list-ports
```

### 步骤 7: 配置腾讯云安全组

1. 登录腾讯云控制台
2. 进入 **云服务器** → **安全组**
3. 选择服务器对应的安全组
4. 点击 **入站规则** → **添加规则**
5. 配置：
   - **类型**: 自定义
   - **来源**: 0.0.0.0/0（或限制特定IP）
   - **协议端口**: TCP:8000
   - **策略**: 允许
6. 保存

### 步骤 8: 启动服务

#### 测试启动

```bash
cd /opt/fastapi-ocr
source .venv/bin/activate
chmod +x start_cloud.sh
./start_cloud.sh
```

#### 使用 systemd 管理服务（推荐）

创建服务文件：

```bash
sudo nano /etc/systemd/system/fastapi-ocr.service
```

内容：

```ini
[Unit]
Description=FastAPI OCR Translation Service
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=/opt/fastapi-ocr
Environment="PATH=/opt/fastapi-ocr/.venv/bin"
ExecStart=/opt/fastapi-ocr/.venv/bin/uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 4
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

启动服务：

```bash
# 重载systemd配置
sudo systemctl daemon-reload

# 启动服务
sudo systemctl start fastapi-ocr

# 设置开机自启
sudo systemctl enable fastapi-ocr

# 查看状态
sudo systemctl status fastapi-ocr

# 查看日志
sudo journalctl -u fastapi-ocr -f
```

## 🔒 安全配置

### 1. 使用 Nginx 反向代理（推荐）

#### 安装 Nginx

```bash
# Ubuntu/Debian
apt-get install -y nginx

# CentOS/RHEL
yum install -y nginx
```

#### 配置 Nginx

```bash
sudo nano /etc/nginx/sites-available/fastapi-ocr
```

内容：

```nginx
server {
    listen 80;
    server_name your_domain.com;  # 或使用服务器IP

    client_max_body_size 50M;  # 允许上传大文件

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

启用配置：

```bash
# Ubuntu/Debian
sudo ln -s /etc/nginx/sites-available/fastapi-ocr /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl restart nginx

# CentOS/RHEL
sudo cp /etc/nginx/sites-available/fastapi-ocr /etc/nginx/conf.d/
sudo nginx -t
sudo systemctl restart nginx
```

### 2. 配置 HTTPS（使用 Let's Encrypt）

```bash
# 安装 Certbot
apt-get install -y certbot python3-certbot-nginx

# 获取证书
sudo certbot --nginx -d your_domain.com

# 自动续期
sudo certbot renew --dry-run
```

### 3. 限制文件上传大小

在 `app/main.py` 中已配置，Nginx 配置中也已设置 `client_max_body_size 50M`

## 📊 监控和维护

### 查看服务状态

```bash
# systemd服务状态
sudo systemctl status fastapi-ocr

# 查看日志
sudo journalctl -u fastapi-ocr -n 100

# 查看进程
ps aux | grep uvicorn

# 查看端口
netstat -tuln | grep 8000
```

### 重启服务

```bash
sudo systemctl restart fastapi-ocr
```

### 更新代码

```bash
cd /opt/fastapi-ocr
source .venv/bin/activate

# 拉取新代码或上传新文件
# ...

# 安装新依赖（如果有）
pip install -r requirements.txt

# 重启服务
sudo systemctl restart fastapi-ocr
```

## 🧪 测试部署

### 1. 测试本地访问

```bash
curl http://127.0.0.1:8000
```

### 2. 测试公网访问

在浏览器中访问：
```
http://your_server_ip:8000
```

### 3. 测试API

```bash
curl -X POST "http://your_server_ip:8000/task/run?from_lang=zh&to_lang=en" \
  -F "file=@test_image.jpg"
```

## 🐛 常见问题

### 问题 1: 无法访问

**检查清单：**
- [ ] 防火墙是否开放端口
- [ ] 腾讯云安全组是否配置
- [ ] 服务是否正常运行
- [ ] IP地址是否正确

### 问题 2: 内存不足

**解决：**
- 减少 workers 数量：`--workers 2`
- 增加服务器内存
- 使用 swap 空间

### 问题 3: 上传文件失败

**检查：**
- Nginx `client_max_body_size` 配置
- FastAPI 文件大小限制
- 磁盘空间是否充足

### 问题 4: OCR识别失败

**检查：**
- PaddleOCR 模型是否正确下载
- 内存是否充足
- 查看日志：`sudo journalctl -u fastapi-ocr -f`

## 📝 性能优化

### 1. 使用 Gunicorn + Uvicorn Workers

```bash
pip install gunicorn

gunicorn app.main:app \
    -w 4 \
    -k uvicorn.workers.UvicornWorker \
    -b 0.0.0.0:8000 \
    --timeout 120
```

### 2. 启用缓存

考虑使用 Redis 缓存OCR结果

### 3. 异步处理大文件

对于大文件，考虑使用 Celery 异步处理

## 🔄 备份和恢复

### 备份

```bash
# 备份数据库
cp data/app.db data/app.db.backup

# 备份配置文件
tar -czf backup-$(date +%Y%m%d).tar.gz .env data/
```

### 恢复

```bash
# 恢复数据库
cp data/app.db.backup data/app.db

# 恢复配置
tar -xzf backup-YYYYMMDD.tar.gz
```

## 📞 技术支持

如遇问题，请检查：
1. 服务日志：`sudo journalctl -u fastapi-ocr -f`
2. Nginx日志：`/var/log/nginx/error.log`
3. 系统资源：`htop` 或 `top`

---

**部署完成后，访问地址：**
- HTTP: `http://your_server_ip:8000`
- HTTPS (如果配置): `https://your_domain.com`

