import requests
import json
import time
import os
from dotenv import load_dotenv
# 加载 .env 文件
load_dotenv()
api_key=os.getenv("API_KEY")
base_url=os.getenv("BASE_URL")
start_time = time.time()
response = requests.get(
  url="https://openrouter.ai/api/v1/key",
  headers={
    "Authorization": f"Bearer {api_key}"
  }
)
end_time = time.time()
latency = end_time - start_time
print(f"响应总耗时: {latency:.2f} 秒")
print(json.dumps(response.json(), indent=2))
