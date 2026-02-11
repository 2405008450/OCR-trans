from pydantic import BaseModel

class TaskOption(BaseModel):
    language: str = "zh"
    max_tokens: int = 2048
