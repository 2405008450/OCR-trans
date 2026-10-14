"""AI 检查相关异常。"""


class AICheckIncompleteError(RuntimeError):
    """AI 未给出有效结果（调用失败 / 返回为空 / 返回被截断 / 无法解析）。

    这类情况不能当作"没有错误"处理，必须让任务失败并提示用户重新检查。
    """
