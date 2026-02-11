# 局域网部署指南

## 📋 部署步骤

### 1. 启动服务器（局域网模式）

#### Windows 系统

**方法一：使用启动脚本（推荐）**
```bash
# 双击运行
start_server.bat
```

**方法二：手动启动**
```bash
# 激活虚拟环境
.venv\Scripts\activate

# 启动服务（绑定到所有网络接口）
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

#### Linux/Mac 系统

**方法一：使用启动脚本**
```bash
chmod +x start_server.sh
./start_server.sh
```

**方法二：手动启动**
```bash
# 激活虚拟环境
source .venv/bin/activate

# 启动服务
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

### 2. 配置防火墙

#### Windows 防火墙配置

1. **打开防火墙设置**
   - 按 `Win + R`，输入 `wf.msc`，回车
   - 或：控制面板 → Windows Defender 防火墙 → 高级设置

2. **添加入站规则**
   - 点击左侧"入站规则" → 右侧"新建规则"
   - 选择"端口" → 下一步
   - 选择"TCP"，特定本地端口：`8000` → 下一步
   - 选择"允许连接" → 下一步
   - 全部勾选（域、专用、公用）→ 下一步
   - 名称：`FastAPI OCR服务` → 完成

3. **或者使用命令行（管理员权限）**
   ```powershell
   New-NetFirewallRule -DisplayName "FastAPI OCR服务" -Direction Inbound -LocalPort 8000 -Protocol TCP -Action Allow
   ```

#### Linux 防火墙配置（iptables）

```bash
# Ubuntu/Debian
sudo ufw allow 8000/tcp

# CentOS/RHEL
sudo firewall-cmd --permanent --add-port=8000/tcp
sudo firewall-cmd --reload
```

### 3. 访问地址

#### 本机访问
- `http://127.0.0.1:8000`
- `http://localhost:8000`

#### 局域网访问
- `http://192.168.31.125:8000` （你的本机IP）

#### 其他设备访问
确保其他设备与服务器在同一局域网（同一WiFi或同一网段），然后使用：
- `http://192.168.31.125:8000`

### 4. 验证部署

#### 检查服务是否启动
```bash
# Windows
netstat -an | findstr :8000

# Linux/Mac
netstat -an | grep :8000
# 或
ss -tuln | grep :8000
```

应该看到类似输出：
```
TCP    0.0.0.0:8000           0.0.0.0:0              LISTENING
```

#### 测试访问
1. 在本机浏览器访问 `http://127.0.0.1:8000`
2. 在其他设备浏览器访问 `http://192.168.31.125:8000`
3. 如果都能访问，说明部署成功！

## 🔧 常见问题

### 问题1：其他设备无法访问

**可能原因：**
- 防火墙未开放端口
- 服务器未绑定到 `0.0.0.0`
- 设备不在同一局域网

**解决方法：**
1. 检查防火墙设置（见上方）
2. 确认启动命令包含 `--host 0.0.0.0`
3. 确保设备连接同一WiFi/网络

### 问题2：端口被占用

**解决方法：**
```bash
# 更改端口（例如改为8001）
uvicorn app.main:app --host 0.0.0.0 --port 8001 --reload
```

然后访问 `http://192.168.31.125:8001`

### 问题3：IP地址变化

如果路由器重启或网络变化，IP地址可能会改变。

**查看当前IP：**
```bash
# Windows
ipconfig

# Linux/Mac
ifconfig
# 或
ip addr show
```

### 问题4：移动设备无法访问

**检查：**
1. 手机/平板是否连接同一WiFi
2. 防火墙是否允许移动设备访问
3. 尝试使用IP地址而不是域名

## 📱 移动设备访问

### iOS/Android
1. 确保设备连接同一WiFi
2. 打开浏览器（Safari/Chrome）
3. 输入：`http://192.168.31.125:8000`
4. 开始使用！

## 🔒 安全建议

### 生产环境部署

如果需要在生产环境部署，建议：

1. **使用HTTPS**
   ```bash
   uvicorn app.main:app --host 0.0.0.0 --port 8000 --ssl-keyfile key.pem --ssl-certfile cert.pem
   ```

2. **添加认证**
   - 在FastAPI中添加API密钥验证
   - 或使用OAuth2

3. **限制访问IP**
   - 在防火墙中设置只允许特定IP访问

4. **使用反向代理**
   - Nginx
   - Apache
   - Caddy

## 📊 性能优化

### 多进程部署
```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 4
```

### 使用Gunicorn（Linux/Mac）
```bash
pip install gunicorn
gunicorn app.main:app -w 4 -k uvicorn.workers.UvicornWorker -b 0.0.0.0:8000
```

## 🎯 快速启动命令

### Windows
```powershell
.venv\Scripts\activate
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

### Linux/Mac
```bash
source .venv/bin/activate
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

---

**提示**：如果IP地址是 `192.168.31.125`，其他设备访问地址为：
```
http://192.168.31.125:8000
```


