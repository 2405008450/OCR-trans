"""隔离质检网络调用，防止供应商内部重试突破整个质检的时间预算。"""
import multiprocessing


def _worker(connection, function, args, kwargs):
    try:
        connection.send((True, function(*args, **kwargs)))
    except BaseException as exc:
        # 不通过任务日志泄漏 URL 中的 key 或原始供应商响应。
        connection.send((False, type(exc).__name__))
    finally:
        connection.close()


def bounded_call(function, *args, budget_seconds, **kwargs):
    if budget_seconds <= 0:
        raise TimeoutError("质检调用预算耗尽")
    context = multiprocessing.get_context("spawn")
    receiver, sender = context.Pipe(duplex=False)
    process = context.Process(target=_worker, args=(sender, function, args, kwargs), daemon=True)
    try:
        process.start()
        sender.close()
        if not receiver.poll(budget_seconds):
            raise TimeoutError("质检调用超时，已终止本次调用")
        success, result = receiver.recv()
        if not success:
            raise RuntimeError(f"质检隔离调用失败：{result}")
        return result
    finally:
        sender.close()
        receiver.close()
        if process.pid is not None:
            if process.is_alive():
                process.terminate()
            process.join(timeout=2)
            if process.is_alive():
                process.kill()
                process.join(timeout=2)
            process.close()
