import os

import requests
import speedtest

def check_proxy():
    test_url = "http://httpbin.org/ip"  # 或者用 "https://api.ipify.org"
    try:
        # 显式传入代理测试
        response = requests.get(test_url, timeout=5)
        print(f"当前出口 IP: {response.json()['origin']}")
        print("代理连接状态: 成功")
    except Exception as e:
        print(f"代理连接失败: {e}")

check_proxy()
# 创建 Speedtest 对象
st = speedtest.Speedtest()

# 获取最佳服务器
st.get_best_server()

# 测试下载速度（bits per second）
download_speed = st.download() / 1_000_000  # 转换为 Mbps
print(f"下载速度: {download_speed:.2f} Mbps")

# 测试上传速度
upload_speed = st.upload() / 1_000_000  # 转换为 Mbps
print(f"上传速度: {upload_speed:.2f} Mbps")

# 查看 ping 延迟
ping = st.results.ping
print(f"延迟: {ping:.2f} ms")
