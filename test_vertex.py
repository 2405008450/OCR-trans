"""Vertex AI 连接测试脚本 — 在服务器上运行: python test_vertex.py"""
import time
import random
from google import genai
from google.genai import types
from google.genai.errors import ClientError

PROJECT_ID = "gen-lang-client-0128671098"
LOCATION = "global"
MODEL_NAME = "gemini-3-flash-preview"

print(f"[1] 创建 Vertex AI Client (project={PROJECT_ID}, location={LOCATION})")
client = genai.Client(
    vertexai=True,
    project=PROJECT_ID,
    location=LOCATION,
)
print("[1] ✅ Client 创建成功")


def generate_with_retry(model, contents, config=None, max_retries=6):
    delay = 2.0
    for attempt in range(max_retries):
        try:
            return client.models.generate_content(
                model=model,
                contents=contents,
                config=config,
            )
        except ClientError as e:
            msg = str(e)
            if "429" in msg or "RESOURCE_EXHAUSTED" in msg:
                if attempt == max_retries - 1:
                    raise
                sleep_s = delay + random.uniform(0, 1.5)
                print(f"    429 限流，等待 {sleep_s:.1f}s 后重试... ({attempt + 1}/{max_retries})")
                time.sleep(sleep_s)
                delay *= 2
            else:
                raise


print("\n[2] 测试文本生成...")
try:
    resp = generate_with_retry(
        model=MODEL_NAME,
        contents="Explain how AI works in a few words",
    )
    print(f"[2] ✅ 文本生成成功:\n    {resp.text[:200]}")
except Exception as e:
    print(f"[2] ❌ 文本生成失败: {type(e).__name__}: {e}")

print("\n[3] 测试带 config 的生成...")
try:
    resp = generate_with_retry(
        model=MODEL_NAME,
        contents="说一句中文测试",
        config=types.GenerateContentConfig(
            system_instruction="你是一个翻译助手",
            temperature=0,
            max_output_tokens=100,
        ),
    )
    print(f"[3] ✅ 带 config 生成成功:\n    {resp.text[:200]}")
except Exception as e:
    print(f"[3] ❌ 带 config 生成失败: {type(e).__name__}: {e}")

print("\n========== 测试完成 ==========")
