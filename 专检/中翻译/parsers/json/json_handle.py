import json


def filter_json_objects(input_path: str, output_path: str) -> int:
    """
    读取 JSON 文件（数组或多个对象），删除"译文数值"与"译文修改建议值"严格一致的对象。
    返回删除的数量。
    """
    with open(input_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, dict):
        data = [data]

    before = len(data)
    filtered = [obj for obj in data if obj.get("译文数值") != obj.get("译文修改建议值")]
    removed = before - len(filtered)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(filtered, f, ensure_ascii=False, indent=2)

    print(f"共 {before} 条，删除 {removed} 条，保留 {len(filtered)} 条 -> {output_path}")
    return removed


if __name__ == "__main__":
    filter_json_objects
    # import sys
    #
    # if len(sys.argv) < 3:
    #     print("用法: python json_handle.py <输入文件> <输出文件>")
    #     sys.exit(1)
    #
    # filter_json_objects(sys.argv[1], sys.argv[2])
