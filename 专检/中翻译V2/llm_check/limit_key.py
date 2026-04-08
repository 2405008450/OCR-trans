import requests
import json
import time
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[3]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.core.config import settings

api_key = settings.OPENROUTER_API_KEY
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
