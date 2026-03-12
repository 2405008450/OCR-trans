import requests
import json
import time
import os
from dotenv import load_dotenv
# 加载项目根目录 .env（与主应用一致）
_project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "..", "..", ".."))
_env_path = os.path.join(_project_root, ".env")
if os.path.isfile(_env_path):
    load_dotenv(_env_path)
api_key = os.getenv("OPENROUTER_API_KEY") or os.getenv("API_KEY")
base_url = os.getenv("OPENROUTER_BASE_URL") or os.getenv("BASE_URL")
print(api_key)
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
