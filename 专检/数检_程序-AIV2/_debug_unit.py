import sys
sys.path.insert(0, r'd:\project\数检_程序-AI')
import re
from normalizer_total import PATTERNS, extract_numbers

text_cn = '2.单位：元 ;单位：万元'

_cn_unit_decl_pat = re.compile(
    r"单位[：:]\s*(百万|千万|百亿|千亿|亿|千|百|万)?\s*(?:美元|欧元|元人民币|元|人民币)",
    re.I
)

print("=== _cn_unit_decl_pat 匹配 ===")
for m in _cn_unit_decl_pat.finditer(text_cn):
    print(f"  match={m.group()!r}  group1={m.group(1)!r}  span={m.span()}")

print()
print("=== currency_cn 匹配 ===")
for m in PATTERNS["currency_cn"].finditer(text_cn):
    print(f"  match={m.group()!r}  span={m.span()}")

print()
print("=== extract_numbers 结果 ===")
print(extract_numbers(text_cn))
