# 翻译数值审校系统 — 技术架构文档

## 项目概述

本系统用于自动审校双语翻译文档（中译英）中的数值一致性问题。支持 DOCX / XLSX / PPTX / PDF 四种格式，核心流程分为三个阶段：**提取** → **检查** → **修订写入**。

最终输出带 Track Changes 修订标记的译文文档，以及 Excel 审校报告。

---

## 整体架构

```
输入
  ├── 原文文件（中文）
  └── 译文文件（英文）  或  对照 Excel

       ↓ 提取阶段
  ┌─────────────────────────────────┐
  │  full_content.py                │  ← DOCX 按阅读顺序全文提取
  │  extract_any.py                 │  ← XLSX / PPTX / PDF 提取适配
  │  header_extractor.py            │  ← 页眉单独提取
  │  footer_extractor.py            │  ← 页脚单独提取
  └─────────────────────────────────┘

       ↓ 检查阶段
  ┌─────────────────────────────────┐
  │  normalizer.py                  │  ← 数值归化（万/亿/million等）
  │  normalizer_total.py            │  ← 扩展归化（货币/日期/分数等）
  │  extract_values.py              │  ← 规则检查（数值提取+对比）
  │  program_check.py               │  ← AI 复核（DeepSeek/Gemini）
  └─────────────────────────────────┘

       ↓ 写回阶段
  ┌─────────────────────────────────┐
  │  numbering_to_static.py         │  ← 自动编号/目录静态化
  │  replace_revision.py            │  ← 多策略定位 + 替换
  │  revision.py                    │  ← Track Changes XML 操作
  │  replace_clean.py               │  ← 文本清洗工具
  └─────────────────────────────────┘

       ↓ 输出
  ├── 修订版译文文档（.docx / .pdf / .xlsx / .pptx）
  └── 审校报告（.xlsx）
```

---

## 运行入口

### main.py

主流程文件，对外暴露 `run()` 函数。

**两种输入模式：**

- **模式A（对照文件）**：传入预先对齐好的原文/译文 Excel（`alignment_path`），适用于 PDF 或已有对照文件的场景
- **模式B（直接提取）**：传入原文文件 + 译文文件（`src_docx_path` + `tgt_docx_path`），程序自动提取并对齐

**五个阶段：**

1. 正文检查（规则 + AI）
2. 页眉/页脚检查（仅 DOCX）
3. 保存 AI 原始结果 JSON（`align_body.json` / `align_body_flat_errors.json` 等）
4. 生成 Excel 审校报告
5. 写入修订（自动编号静态化 → 目录静态化 → 多策略替换）

**关键设计决策：**
- 正文提取过滤掉 `source="header/footer"` 片段，页眉/页脚走独立检查通道，防止区域错乱
- 写回前先调用 `convert_numbering_to_static` + `convert_toc_to_static`，确保自动编号和目录域字段已展开为静态文本，替换定位可靠

---

## 提取层

### full_content.py

DOCX 文档全文提取核心。直接解析 ZIP 内的 `word/document.xml`，不依赖 python-docx。

**提取内容：**
- 正文段落（body）
- 表格单元格（table），附带同行所有单元格的 `row_context`
- 页眉/页脚（header/footer）
- 脚注/尾注（footnote/endnote），内联插入到引用段落后
- 图表文字（chart，解析 chart XML 的 `<a:t>` 和 `<c:v>`）
- 浮动文本框（textbox）

**自动编号支持：**

Word 自动编号存储在 `numbering.xml`，不存在于 `<w:t>` 中。`full_content.py` 通过以下步骤还原：

1. 解析 `numbering.xml`：读取 `<w:abstractNum>`（抽象编号定义）和 `<w:num>`（实例化）
2. 解析 `styles.xml`：读取样式（如 `heading 1/2/3`）绑定的编号实例
3. 遍历段落时检查 `<w:numPr>`，优先读段落自身，无则从样式继承
4. `NumberingState.advance()` 推进计数、重置更深级别、展开 `lvlText` 模板

支持的 `numFmt`：`decimal`、`chineseCounting`、`japaneseCounting`、`chineseLegalSimplified`、`upperLetter`/`lowerLetter`、`upperRoman`/`lowerRoman`、`decimalFullWidth` 等。

**返回结构：**

```python
@dataclass
class TextSegment:
    source: str        # body / header / footer / footnote / endnote / chart / table / textbox
    text: str
    xml_path: str      # ZIP 内 XML 路径（写回用）
    para_index: int    # 段落顺序索引（夹逼定位用）
    row_context: str   # 表格行上下文（同行所有单元格 tab 拼接）
```

### extract_any.py

多格式统一提取适配器，支持 XLSX / PPTX / PDF。对外暴露 `extract()` 接口，返回与 `TextSegment` 兼容的片段列表。

### header_extractor.py / footer_extractor.py

分别提取 DOCX 文档所有节的页眉/页脚文本，返回字符串列表。与正文提取完全隔离，检查结果只用 `region="header/footer"` 写回，不会误替换正文。

### body_extractor.py

提取正文纯文本的简化版本（早期版本，现主要由 `full_content.py` 替代）。

---

## 检查层

### normalizer.py

数值归化引擎，将不同表示形式的数值归一化为可比较的浮点数字符串。

**支持的归化规则：**

| 输入 | 归化结果 |
|---|---|
| `14.06万` | `140600.0` |
| `340.368 million` | `340368000.0` |
| `三百二十一` | `321` |
| `XIV` | `14` |
| `January` | `1` |
| `二〇二五年三月` | `2025-03` |

**策略开关：** `chinese_upper`、`chinese_trad`、`month_name`、`english_number`、`roman`，可按需开关。

### normalizer_total.py

扩展版归化，额外支持货币符号（`$`/`¥`/`€`）、百分比、分数、季度（`Q1`/第一季度）、比率等更复杂的规则。

### extract_values.py

规则检查入口，提供三个核心函数：

- `extract_numbers(text)` — 从文本提取数值列表
- `check_row(src, tgt)` — 逐行规则检查，返回是否错误及错误类型
- `run_number_check(alignment_path)` — 从对照 Excel 批量检查
- `check_text_pairs(pairs)` — 从文本对列表批量检查
- `merge_ai_results(rows, ai_map)` — 将 AI 结果合并进规则检查结果

### program_check.py

AI 复核模块，调用大模型（默认 DeepSeek V4 Pro）对规则检查报错的行做二次判断。

**输入：** 每批 `(seq, row)` 列表，发送格式：
```
[0] 原文: ...
[0] 译文: ...
[1] 原文: ...
[1] 译文: ...
```

**输出 Schema（`_ERROR_SCHEMA`）：**
```json
{
  "原文上下文": "...",
  "译文上下文": "...",
  "原文数值": "...",
  "译文数值": "...",
  "替换锚点": "需要被替换的精确片段",
  "译文修改建议值": "替换后的内容",
  "is_source_consistent": false,
  "错误类型": "数值错误",
  "修改理由": "...",
  "违反的规则": "..."
}
```

`is_source_consistent=true` 表示译文忠实还原了原文但原文本身有问题，不生成修订任务。

---

## 写回层

### numbering_to_static.py

修订写入前的预处理，将 Word 动态内容转为静态文本。

- `has_auto_numbering(path)` — 检测文档是否有自动编号
- `convert_numbering_to_static(path)` — 将 `<w:numPr>` 编号展开写入 `<w:t>`
- `convert_toc_to_static(path)` — 将 TOC 域字段展开为静态段落文本（含页码）

**为何必须静态化：** 自动编号数字不存在于 `<w:t>` 中，替换时搜不到；TOC 域字段页码在 `build_para_cache` 中被跳过，夹逼定位找不到前后句。

### replace_revision.py

替换定位的核心，提供 `replace_and_revise_in_docx()` 函数，按优先级依次尝试多种定位策略。

**策略顺序（段落定位）：**

| 策略 | 使用字段 | 说明 |
|---|---|---|
| 0 — 夹逼定位 | `prev_tgt` + `next_tgt` | 用前后句在全文找位置，把含 `old_value` 的候选夹在中间；前后句带页码时自动去掉末尾数字降级匹配 |
| 0A — 自动编号替换 | `old_value` 模式 + `context` | 仅未静态化时触发，处理 `iv.`→`v.` 类编号变更 |
| 0B — 中文编号警告 | `old_value` 模式 | 检测疑似未静态化的中文编号 |
| 1 — 显式锚点匹配 | `anchor_text` + `context` | 锚点命中段落，用 context 相似度过滤；短锚点（<15字符）且相似度<0.1 的误命中被过滤 |
| 2 — 上下文锚点匹配 | `context` | 从 context 提取含 old_value 的子串作锚点 |
| 3~6 — 候选打分 | `old_value` + `context` | 单次全文遍历，从严格到宽松四档正则，context 相似度打分排序 |
| 7 — 脚注/尾注 | `old_value` | 正文未找到时检查脚注区域 |
| 兜底 — 批注 | 全文含 `old_value` 的段落 | 无法精确定位时全部标注批注 |

**夹逼定位细节：**
- `_find_positions()` 精确匹配 → 去掉末尾数字后匹配（处理目录页码）→ 相似度 ≥ 0.85 兜底
- 夹逼条件：`prev_pos < candidate_pos < next_pos`，且前后间距 ≤ 50 段
- 多个夹逼命中时取物理距离最近的；距离相同则拒绝替换

**段落内替换策略（`_execute_replace`）：**

| 顺序 | 方式 |
|---|---|
| 1 | 精确子串匹配 |
| 2 | 清洗后匹配（去零宽字符、特殊空格） |
| 3 | 编号归一化（全半角、空格） |
| 4 | 中文后缀剥离（年/月/万/亿等） |
| 5 | sym 符号容忍（×/·等符号位置宽松匹配） |
| 6 | 无空格单 run 匹配 |

### revision.py

Word Track Changes XML 操作层，提供 `RevisionManager`。

- `replace_in_paragraph(para, old, new)` — 在段落内用 `<w:del>` + `<w:ins>` 包裹替换，支持跨 run 场景
- `replace_run_text(run, new)` — 单 run 替换
- `insert_comment(para, text)` — 在段落插入批注（兜底策略用）

`replace_in_paragraph` 的 run 处理逻辑：
1. 拼接所有 run 文本，`find()` 定位 `old_text` 起止位置
2. 找出覆盖该位置的所有 run（单 run 或多 run）
3. 单 run：拆分为 prefix + `<w:del>old</w:del><w:ins>new</w:ins>` + suffix
4. 多 run：删除所有涉及 run，整体重建

### replace_clean.py

文本清洗工具集，提供 `clean_text_thoroughly()`，统一处理智能引号、全半角、零宽字符、多余空格等。

### clean_replace_duplicates.py

替换建议值去重，检测 `new_val` 与上下文前后的重叠片段并剪裁，避免双写。例如上下文已包含前缀，建议值不需要重复带上前缀。

---

## 工具模块

### apply_from_result.py

独立的写回脚本，从已保存的 `align_body_errors.json` 直接执行修订，无需重新调用大模型。适用于重跑写回、调整策略后重试的场景。

### test_quick_replace.py

快速测试工具，从 `align_body_flat_errors.json` 加载错误列表，对指定文档执行替换并统计成功率。支持 DOCX / XLSX / PPTX / PDF。

### test_quick_replace_number.py / test_single_replace.py

单条替换调试工具，用于验证特定错误的定位和替换效果。

### backup_copy/backup_manager.py

文件备份管理，写回前自动创建带时间戳的备份副本，所有写回操作在副本上执行，不破坏原文件。

---

## 格式支持子模块

### excel/

- `excel_parser.py` — XLSX 文本提取
- `excel_replacer.py` — XLSX 批注式替换（高亮 + 批注，不支持 Track Changes）
- `clean_json.py` — Excel 数据清洗工具

### ppt/

- `pptx_parser.py` — PPTX 文本提取（遍历所有 slide 的文本框）
- `pptx_replacer.py` — PPTX 批注式替换

### pdf/

- `pdf_parser.py` — PDF 文本提取（基于 pdfminer）
- `pdf_replacer_improved.py` — PDF 文本覆盖 + 批注替换（基于 PyMuPDF）
- `pdf_annotator.py` — PDF 批注操作
- `numbering_replacer.py` — PDF 编号特殊处理
- `text_matcher.py` — PDF 文本定位匹配

### word/

- `doc_parser.py` — 旧版 `.doc` 格式解析（通过 LibreOffice 转换后处理）

---

## 数值检查子模块 num_checker/

独立的高级数值检查模块，包含机器学习方法：

- `checker.py` — 主入口
- `_parser_core.py` — 数值解析核心
- `alignment_matrix.py` — 原译文数值对齐矩阵
- `crf_discriminator.py` — CRF 序列标注判别器
- `rl_discriminator.py` — 强化学习判别器
- `domain_rules.py` — 领域规则（金融/医疗/工业等）
- `learning_pipeline.py` — 自学习流水线，从历史错误中更新模型
- `symbolic_parser.py` — 符号化解析（分数、比率、范围等）

---

## NER 子模块

基于 PaddleNLP 的命名实体识别模块，用于双语实体对齐辅助。

- `SimAlign.py` — 基于 SimAlign 的词对齐
- `demo.py` / `demo1.py` — 使用示例

---

## 数据流

### 模式B（直接提取，最完整路径）

```
原文 DOCX → full_content.py → src_segs（过滤 header/footer）
译文 DOCX → full_content.py → tgt_segs（过滤 header/footer）
                                    ↓ 段落级精确对齐（para_index）
                              pairs[(src_text, tgt_text, para_index, row_context)]
                                    ↓
                         check_text_pairs → body_rows（规则检查结果）
                                    ↓ 仅规则错误行
                         _llm_check_block → body_ai_map（AI 检查结果）
                                    ↓
                         merge_ai_results → body_final
                                    ↓
              _ai_map_to_flat_errors → align_body_flat_errors.json
                （每条 error 附带 prev_tgt/next_tgt，供夹逼定位用）
                                    ↓
              numbering_to_static + convert_toc_to_static（预处理）
                                    ↓
              replace_and_revise_in_docx（多策略替换）
                                    ↓
              修订版译文文档 + 审校报告 Excel
```

### 夹逼定位数据流

```
final_rows[idx]["译文"]     → context（AI 返回的含错误数值的句子片段）
final_rows[idx-1]["译文"]   → prev_tgt（上一行完整译文）
final_rows[idx+1]["译文"]   → next_tgt（下一行完整译文）

_locate_by_neighbor_sentences():
  1. prev_tgt → 在 para_cache 搜 prev_positions（带页码时去末尾数字降级）
  2. next_tgt → 在 para_cache 搜 next_positions（同上）
  3. old_value → 在 para_cache 搜 candidate_positions
  4. 夹逼：prev_pos < candidate < next_pos，间距 ≤ 50
  5. 唯一命中 → _execute_replace
     多命中 → 取物理距离最近的；距离相同 → 拒绝
```

---

## 中间产物文件（output/）

| 文件 | 内容 |
|---|---|
| `align_body.json` | 全部行的规则+AI检查结果（含 seq/原文/译文） |
| `align_body_errors.json` | 仅含 errors 数组的行（供 apply_from_result.py 使用） |
| `align_body_flat_errors.json` | 每条 error 展开为独立条目，附 prev_tgt/next_tgt（供 test_quick_replace.py 使用） |
| `align_header.json` | 页眉检查结果 |
| `align_footer.json` | 页脚检查结果 |
| `*_alignment.json` | 原文/译文对照 JSON |
| `*_alignment.xlsx` | 原文/译文对照 Excel（可复用作模式A输入） |
| `*_output.xlsx` | 最终审校报告 Excel |

---

## 配置

`.env` 中配置大模型 API Key（`OPENROUTER_API_KEY` 或直连 DeepSeek/Gemini Key）。

`.env.example` 提供配置模板。

---

## 依赖

- `python-docx` — DOCX 读写
- `lxml` — XML 底层解析
- `pandas` / `openpyxl` — Excel 读写
- `PyMuPDF (fitz)` — PDF 处理
- `pdfminer.six` — PDF 文本提取
- `python-pptx` — PPTX 处理
- `openai` — 大模型 API 调用（兼容 OpenRouter）
- `python-dotenv` — 环境变量加载
