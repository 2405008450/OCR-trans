class MCPRegistry:
    def __init__(self):
        # 存储工具的 JSON 定义
        self.tools_metadata = []
        # 存储工具名到函数的映射
        self.handlers = {}

    def register(self, name, description, input_schema):
        """装饰器：注册一个 MCP 工具"""
        def decorator(func):
            # 1. 保存元数据
            self.tools_metadata.append({
                "name": name,
                "description": description,
                "inputSchema": input_schema
            })
            # 2. 绑定处理函数
            self.handlers[name] = func
            return func
        return decorator
# 实例化注册器
registry = MCPRegistry()