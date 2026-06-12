# num_checker 自学习闭环改造方案

## 1. 当前现状

项目当前已经具备以下能力：

- 规则/符号提取：`symbolic_parser.py`
- 结果对齐校验：`alignment_matrix.py`
- 可训练的多义词判别器：`crf_discriminator.py`
- 独立训练入口：`train_crf.py`

项目当前缺失的关键能力：

- 反馈采集：运行结果没有沉淀为可复用样本
- 人工审核：无法把"真实翻译错误"和"模型误判"分开
- 数据治理：没有统一的学习样本存储格式
- 训练门禁：没有新旧模型对比和自动准入条件
- 版本切换：没有候选模型与正式模型的切换策略

因此，当前系统属于"可离线再训练的检查器"，还不是"真正自学习的闭环系统"。

## 2. 改造目标

把项目升级为一个最小可用的"人机协同自学习系统"：

1. 运行检查任务时，自动沉淀可审核案例
2. 人工在统一数据文件中补充审核结论和训练信号
3. 系统基于审核通过的数据增量训练 CRF
4. 新模型必须通过评估门禁才替换正式模型
5. 全流程尽量复用现有解析、对齐、CRF 代码

## 3. 目标架构

```text
                +----------------------+
                |  bilingual excel     |
                |  or demo sentence    |
                +----------+-----------+
                           |
                           v
                +----------------------+
                |   checker.check_*    |
                | parse -> filter ->   |
                | align -> report      |
                +----------+-----------+
                           |
                           v
                +----------------------+
                | feedback collector   |
                | queue case to store  |
                +----------+-----------+
                           |
                           v
                +----------------------+
                | learning_store.json  |
                | pending / approved / |
                | rejected cases       |
                +----------+-----------+
                           |
                      human review
                           |
                           v
                +----------------------+
                | training pipeline    |
                | build feedback set   |
                | split train / eval   |
                +----------+-----------+
                           |
                           v
                +----------------------+
                | candidate CRF model  |
                +----------+-----------+
                           |
                           v
                +----------------------+
                | evaluation gate      |
                | compare baseline     |
                | vs candidate         |
                +----------+-----------+
                           |
            +--------------+---------------+
            |                              |
            v                              v
 +----------------------+      +----------------------+
 | promote active model |      | keep baseline model  |
 | overwrite crf_model  |      | and output report    |
 +----------------------+      +----------------------+
```

## 4. 模块拆分

### 4.1 反馈采集模块

新增模块：`learning_pipeline.py`

职责：

- 定义统一学习样本存储格式
- 将检查结果写入 `learning_data/learning_store.json`
- 对重复案例去重并累计出现次数
- 允许审核状态维护为 `pending/approved/rejected`

### 4.2 审核数据层

学习样本统一结构：

```json
{
  "case_id": "sha1...",
  "src_text": "...",
  "tgt_text": "...",
  "src_values": ["3", "2024"],
  "tgt_values": ["2", "2024"],
  "errors": [{"error_type": "MISSING", "message": "..."}],
  "is_correct": false,
  "summary": "...",
  "review_status": "pending",
  "review_notes": "",
  "training_signal": {
    "text": "",
    "target_word": "",
    "is_numeric": null
  },
  "seen_count": 1
}
```

人工审核时不再要求手工从零填写 `training_signal`。

系统在案例入库时会自动预填：

- `training_signal`
- `candidate_signals`
- `model_version`

因此审核人员的主要动作变成：

- 确认这是模型误判还是翻译真实错误
- 若是模型误判，只需确认 `是数字 / 不是数字`
- 若是真实翻译错误，标记为 `rejected`

### 4.3 训练流水线

训练时使用两部分数据：

- 原有 `SEED_DATA`
- 审核通过的反馈样本

执行流程：

1. 从学习仓库读取 `approved` 样本
2. 过滤出 `training_signal` 完整的样本
3. 将反馈样本分成训练集和评估集
4. 与种子数据合并训练新模型
5. 对反馈样本做配比控制，避免近期错误案例过度主导训练
6. 在固定评估集上比较新旧模型

### 4.4 准入门禁

候选模型必须同时满足：

- 反馈综合评估不低于基线模型
- 种子测试集准确率不低于基线模型

只有通过门禁，才允许覆盖正式模型 `crf_model.pkl`。

同时维护版本锚点：

- `model_meta.json` 中保存当前正式模型版本
- `learning_store.json` 中每个案例记录 `model_version`
- 若旧版本报出的错误在新版本中不再复现，可作为后续归档依据

### 4.5 运行入口

改造 `run.py`：

- 支持检查任务结束后自动入队学习样本
- 默认仅入队错误案例
- 可选入队全部案例

新增训练入口：

- `learning_pipeline.py` 提供训练与准入函数

## 5. 为什么不直接做全自动在线学习

当前项目是质量敏感型检测系统，直接做"运行后自动训练并自动上线"风险很高：

- 错误样本中混有真实翻译错误，不等于模型错误
- 自动生成标签噪声过大，容易污染模型
- 没有稳定基准集时，自我强化可能越学越偏

因此本次改造采用"人工审核在环"方案，这是最稳妥且最容易落地的自学习路径。

## 6. 审核方式

为避免直接编辑 JSON 造成格式损坏，新增交互式 CLI 审核工具：

```bash
python -m num_checker.reviewer --loop
```

审核员每次只需选择：

- `y`：模型误判，应判为数字
- `n`：模型误判，应判为非数字
- `t`：真实翻译错误，拒绝进入训练
- `r`：其他原因拒绝训练

若标记为 `t`，该案例会被视为业务错题，而不是模型训练样本。

## 7. 本次落地范围

本次代码改造只做最小闭环，不做重量级平台化：

- 做：反馈入库、审核格式、增量训练、门禁准入、CLI 接入
- 不做：Web 审核台、数据库、任务调度器、模型注册中心

## 8. 使用方式

### 7.1 采集案例

```bash
python -m num_checker.run --input your.xlsx --queue-feedback
```

默认把错误案例写入：

`num_checker/learning_data/learning_store.json`

### 8.2 人工审核

推荐使用交互式审核器，而不是手改 JSON：

```bash
python -m num_checker.reviewer --loop
```

如果确实需要直接编辑仓库，重点只需修改：

- `review_status`
- `review_label`
- `training_signal.is_numeric`
- `review_notes`

### 8.3 训练并尝试升级模型

在项目根目录 `D:\project\数检_程序-AI` 下运行：

```bash
python -m num_checker.learning_pipeline --train --seed-min-share 0.5
```

如果当前就在包目录 `D:\project\数检_程序-AI\num_checker` 下运行：

```bash
python learning_pipeline.py --train --seed-min-share 0.5
```

若候选模型通过门禁，会自动覆盖正式模型。

### 8.4 导出翻译错题本

```bash
python -m num_checker.learning_pipeline --export-translation-errors
```

该导出只包含被审核为 `translation_error` 的 `rejected` 样本。

## 9. 后续可扩展路线

后续可以继续增强：

- 审核界面：把 JSON 文件审核升级为可视化界面
- 主动学习：优先挑低置信度样本送审
- 自动任务调度：定期训练候选模型
- 模型版本管理：保留历史模型与评估快照
- 多模型策略：规则、CRF、LLM 投票融合
