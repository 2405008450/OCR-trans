"""独立环境中批量识别图片，一次加载模型，逐文件保存结果。"""
import json
import sys
import hashlib
import importlib.metadata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app.core.config import settings
from app.service import paddle_ocr_service as paddle


def main():
    request = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    settings.PADDLEOCR_ENABLED = True
    settings.PADDLEOCR_PYTHON = ""
    settings.PADDLEOCR_CPU_THREADS = 8
    settings.PADDLEOCR_DPI = 300
    settings.PADDLEOCR_CACHE_DIR = str(Path(request["output_dir"]) / "image_cache")
    # 仅本次批量任务启用经本机验证的CPU加速，失败时退回兼容模式。
    original_configuration = paddle._configuration
    fast = True
    def configuration():
        config = original_configuration()
        config["enable_mkldnn"] = fast
        return config
    paddle._configuration = configuration
    # 修正后、兼容模式产生的结果也可续跑，避免再识别已完成图片。
    versions = {name:importlib.metadata.version(name) for name in ('paddleocr','paddlex','PyMuPDF')}
    signatures = {hashlib.sha256(json.dumps({'schema':4,'config':config,'dpi':300,
        'min_score':settings.PADDLEOCR_MIN_SCORE,'versions':versions},sort_keys=True).encode()).hexdigest()[:20]
        for config in (original_configuration(), configuration())}
    original = paddle._recognize_page

    def native_resolution(page, engine, dpi):
        images = page.get_image_info()
        if images:
            info = images[0]
            dpi = min(dpi, 72 * info["width"] / page.rect.width, 72 * info["height"] / page.rect.height)
        return original(page, engine, dpi)

    paddle._recognize_page = native_resolution
    target = Path(request["output_dir"]) / "image_results"
    target.mkdir(exist_ok=True)
    for index, source in enumerate(request["images"], 1):
        try:
            saved = target / f"{source['index']:03d}.json"
            result = json.loads(saved.read_text(encoding='utf-8')) if saved.exists() else None
            if result and (result.get('failed_pages') or Path(result.get('cache_directory','')).name not in signatures):
                result = None
            if result is None:
                result = paddle.extract_paddle_plain_text(file_path=source["path"])
                if result.get('failed_pages') and fast:
                    fast = False
                    result = paddle.extract_paddle_plain_text(file_path=source["path"])
        except Exception as exc:
            result = {"text": "", "page_results": [], "failed_pages": [1], "error": str(exc)}
        (target / f"{source['index']:03d}.json").write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
        print(json.dumps({"index": source["index"], "done": index, "total": len(request["images"]),
                          "path": source["path"], "failed": result.get("failed_pages", [])}, ensure_ascii=True), flush=True)


if __name__ == "__main__":
    main()
