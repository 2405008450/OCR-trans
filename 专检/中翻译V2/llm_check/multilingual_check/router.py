import json
import os
import sys
from pathlib import Path
from openai import OpenAI

PROJECT_ROOT = Path(__file__).resolve().parents[4]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.core.config import settings

client = OpenAI(
    api_key=settings.OPENROUTER_API_KEY,
    base_url=settings.OPENROUTER_BASE_URL,
)

def ai_router(user_query, tools_metadata):
    """
    这是一个逻辑演示。在实际中，你会把 user_query 和 tools_metadata
    发给一个 LLM（如本地的 Ollama 或云端模型）。
    """
    # 构造 Prompt 给 AI
    system_prompt = f"""
    你是一个工具调度员。
    当前可用工具如下：{json.dumps(tools_metadata)}
    用户说："{user_query}"

    请只返回 JSON 格式：{{"tool_name": "工具名", "arguments": {{...}}}}
    如果不匹配任何工具，请返回 {{"tool_name": "claude", "arguments": {{"prompt": "{user_query}"}}}}
    确保参数名称与工具定义中的 inputSchema 完全一致
    """
    try:
        response = client.chat.completions.create(
            model=os.getenv("ZHONGFANYI_MODEL_NAME", "google/gemini-3.1-pro-preview"),
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_query}
            ],
            response_format={"type": "json_object"},  # 强制返回 JSON 格式
            temperature=0.1  # 降低随机性，保证稳定性
        )

        # 解析 AI 返回的内容
        decision = json.loads(response.choices[0].message.content)
        
        # 兼容模型返回 list 的情况，如 [{"tool_name": ...}]
        if isinstance(decision, list) and len(decision) > 0:
            decision = decision[0]
        
        # 兼容模型返回 {"name": ...} 而非 {"tool_name": ...} 的情况
        if "name" in decision and "tool_name" not in decision:
            decision["tool_name"] = decision.pop("name")
        
        return decision

    except Exception as e:
        print(f"[*] AI 路由决策出错: {e}")
        # 兜底方案：如果 AI 挂了，默认尝试交给通用模型处理
        return {"tool_name": "claude", "arguments": {"prompt": user_query}}

    # 模拟 AI 返回的结果
    # 实际开发中，这里调用 client.chat.completions.create(...)
    # if "天气" in user_query:
    #     return {"name": "get_weather", "arguments": {"city": "Shanghai"}}
    # elif "时间" in user_query:
    #     return {"name": "get_time", "arguments": {}}
    # else:
    #     return {"name": "claude", "arguments": {"prompt": user_query}}
