# PaddleOCR 辅助字数统计

字数统计页面的 OCR 模型列表增加 `local/paddleocr`。服务器启用 PaddleOCR 时，页面和未指定模型的请求默认使用本地识别；未启用时沿用现有 LLM 默认模型，用户仍可手动选择其他模型。共享目录必须将 OCR 模式设为“开启”；“智能”仍只对单文件自动识别。PaddleOCR 不调用 LLM、不转换 DOCX，统计沿用现有 Word 近似口径。

## 安装

推荐使用独立环境，避免 OCR 依赖影响 Web 服务。在项目根目录运行：

```powershell
python -m venv .venv-paddleocr
.\.venv-paddleocr\Scripts\python.exe -m pip install -r requirements-paddleocr.txt PyMuPDF==1.27.2 pydantic-settings python-dotenv
```

GPU 环境不安装 requirements-paddleocr.txt 中的 CPU paddlepaddle 包；根据 PaddlePaddle 官方兼容矩阵安装 paddlepaddle-gpu，以及这里固定的 paddleocr 版本。不要同时安装 CPU 与 GPU 引擎。

在服务器 `.env` 配置（路径按实际环境修改），再重启服务：

```dotenv
PADDLEOCR_ENABLED=true
PADDLEOCR_PYTHON=E:/fastapi-llm-demo/.venv-paddleocr/Scripts/python.exe
PADDLEOCR_DEVICE=cpu
PADDLEOCR_CPU_THREADS=4
PADDLEOCR_DPI=200
PADDLEOCR_MIN_SCORE=0.6
PADDLEOCR_TIMEOUT_SECONDS=1800
PADDLEOCR_CACHE_DIR=E:/fastapi-llm-demo/data/paddleocr_cache
```

默认组合为 PP-OCRv5_mobile_det 和 latin_PP-OCRv5_mobile_rec，针对CPU服务器及英语/葡萄牙语等拉丁系文字。可将 `PADDLEOCR_DETECTION_MODEL=PP-OCRv5_server_det` 切换到较大的检测模型。中文资料需要配置支持中文的识别模型；改变模型和 DPI 会产生新的缓存。GPU 设置 `PADDLEOCR_DEVICE=gpu:0`，需实际验证 CUDA 和显存。模型首次运行会下载，离线部署应先预热模型缓存。

独立子进程每文件初始化模型，单文件内复用；同一 Web 进程的 OCR 调用串行。多 Web worker 会分别启动识别进程，服务器资源有限时保持一个 Web worker。每块最多1600像素，保留200像素重叠，按位置消除重复。成功页按文件SHA256、模型配置、包版本缓存；失败页不缓存。服务会保存合并文本，缓存JSON保存逐行文字、置信度、位置及耗时。子进程超时后的成功缓存也可继续复用。

## 报价复核边界

- 复用现有筛页规则：无文字页、少字且大图片页进行整页 OCR；其他页直接提取文字。
- 选中页用整页 OCR 结果替换文字层，避免文字层与识别结果重复计数；失败时保留已有文字并将文件标为失败，现有统计流程不计入合计。
- 默认低于0.6置信度的识别行不计入候选数量；原始内容及被排除行保存在缓存JSON中。低于0.85置信度或没有识别结果的页显示复核提示；没有结果不自动断定空白。置信度不是准确率，人工复核还需检查未检测出的文字。
- 已有大量文字的混合页可能仍遗漏图片文字。工程图微小字、90度文字、分块边缘及识别框冲突也可能影响结果，当前版本没有承诺全量图片区域补充识别。
- 所有PaddleOCR文件附机器统计提示；数字、型号、重复文件及页眉页脚不自动剔除。报告为候选统计，报价前按复核页与代表性样本检查。

## 验证

```powershell
python -m pytest tests/test_paddle_ocr_service.py tests/test_word_count_service.py -q
```

服务器首次启用时，先提交一页已知扫描样本，验证文字、失败状态、缓存JSON和结果下载，再扩大目录范围。未安装模型依赖时，Web服务仍可运行，识别任务会明确失败，不会回退到LLM。

官方文档：https://www.paddleocr.ai/main/en/version3.x/pipeline_usage/OCR.html
