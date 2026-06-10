import requests
import json
import time
import os
from dotenv import load_dotenv
from llm_check.openrouter_config import resolve_openrouter_config
# 加载 .env 文件
load_dotenv()
api_key, base_url = resolve_openrouter_config()
# print(api_key)
start_time = time.time()
response = requests.get(
  url=f"{base_url.rstrip('/')}/key",
  headers={
    "Authorization": f"Bearer {api_key}"
  }
)
end_time = time.time()
latency = end_time - start_time
print(f"响应总耗时: {latency:.2f} 秒")
print(json.dumps(response.json(), indent=2))
