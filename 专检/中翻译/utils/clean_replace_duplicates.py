import difflib
import re


def ultra_safe_replace_v2_debug(context, anchor, suggestion):
    print(f"\n[Debug] 原始建议值: '{suggestion}'")
    if anchor not in context:
        return context, "找不到锚点"

    start_idx = context.find(anchor)
    end_idx = start_idx + len(anchor)
    actual_sug = suggestion

    # --- 前部去重检测 ---
    prefix_window = context[max(0, start_idx - len(suggestion)):start_idx]
    s = difflib.SequenceMatcher(None, prefix_window, suggestion)
    match = s.find_longest_match(0, len(prefix_window), 0, len(suggestion))

    if match.size > 0 and (match.a + match.size == len(prefix_window)) and (match.b == 0):
        actual_sug = suggestion[match.size:]
        print(f"[Debug] 检测到前部重叠: '{suggestion[:match.size]}', 剪裁后: '{actual_sug}'")

    # --- 后部去重检测 ---
    after_window = context[end_idx: end_idx + len(actual_sug)]
    s = difflib.SequenceMatcher(None, actual_sug, after_window)
    match = s.find_longest_match(0, len(actual_sug), 0, len(after_window))

    if match.size > 0 and (match.a + match.size == len(actual_sug)) and (match.b == 0):
        overlap = actual_sug[match.a:]
        actual_sug = actual_sug[:match.a]
        print(f"[Debug] 检测到后部重叠: '{overlap}', 剪裁后: '{actual_sug}'")

    # --- 最终组合 ---
    new_context = context[:start_idx] + actual_sug + context[end_idx:]
    print(f"[Debug] 最终替换内容: '{actual_sug}'")

    return new_context.replace("  ", " "), "OK"

# --- 测试场景 1：单位补偿 ---
ctx1 = "To distribute a cash dividend of RMB6 (tax-inclusive) per ten shares"
anc1 = "6"
sug1 = "RMB4"
res1, _ = ultra_safe_replace_v2_debug(ctx1, anc1, sug1)

print(res1)
# --- 测试场景 2：防止单位重复 ---
ctx2 = "To distribute a cash dividend of RMB6 million per shares"
anc2 = "6"
sug2 = "RMB4"
res2, _ = ultra_safe_replace_v2_debug(ctx2, anc2, sug2)

print(res2)
# --- 测试场景 3：防止文本重复(末尾，开头，中间) ---
ctx3 = "Analysis of Core Business,IV.To distribute a cash, "
anc3 = "IV."
sug3 = "Business,I=V.To"
res3, _ = ultra_safe_replace_v2_debug(ctx3, anc3, sug3)
print(res3)
