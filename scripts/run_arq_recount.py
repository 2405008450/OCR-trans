"""目录重统计：复用经路径和哈希校验的PDF候选统计，重新解析其他文件。"""
import argparse
import csv
import hashlib
import json
import subprocess
import sys
import uuid
from collections import Counter
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app.core.config import settings
from app.db.database import SessionLocal
from app.repository import task_repo
from app.service import word_count_service as service


def read_csv(path):
    with Path(path).open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def normalized(path):
    return str(path).replace("/", "\\").rstrip("\\").casefold()


def sha(path):
    value = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True)
    parser.add_argument("--source-root", default="")
    parser.add_argument("--files", required=True)
    parser.add_argument("--pages", required=True)
    parser.add_argument("--resume-output", default="")
    args = parser.parse_args()
    root = Path(args.root)
    display_root = Path(args.source_root) if args.source_root else root
    original_path = lambda path: display_root / path.relative_to(root)
    imported = read_csv(args.files)
    pages = read_csv(args.pages)
    lookup = {normalized(row["原共享路径"]): row for row in imported}
    if len(lookup) != len(imported):
        raise ValueError("导入清单包含重复路径")
    source_pages = {}
    for page in pages:
        source_pages.setdefault(normalized(page["原共享路径"]), []).append(page)
    candidates = sorted((p for p in root.rglob("*") if p.is_file()), key=lambda p: str(p).casefold())
    if not root.is_dir():
        raise FileNotFoundError(root)
    task_id = str(uuid.uuid4())
    started = datetime.now()
    resume_output = Path(args.resume_output) if args.resume_output else None
    previous = json.loads((resume_output / "执行进度.json").read_text(encoding="utf-8")) if resume_output else None
    if previous:
        task_id, display_no = previous["task_id"], previous["display_no"]
    else:
        with SessionLocal() as db:
            task = task_repo.create_task(db, task_id=task_id, task_type="word_count",
                filename="ARQUITECTURA重统计（复用PDF GPU结果）", status="running", progress=1,
                message="校验PDF结果并重新统计其他文件",
                params_json=json.dumps({"pdf_reuse": str(args.files), "directory_path": str(display_root)}, ensure_ascii=False),
                input_files_json=json.dumps({"directory_path": str(display_root)}, ensure_ascii=False))
            display_no = task.display_no
    output = Path(settings.OUTPUT_DIR) / "word_count" / display_no
    output.mkdir(parents=True, exist_ok=True)
    converted = output / "converted_inputs"
    converted.mkdir(exist_ok=True)
    progress_file = output / "执行进度.json"
    results, sources = [], []
    if resume_output:
        saved = json.loads((resume_output / "文件结果检查点.json").read_text(encoding="utf-8"))
        results = saved if isinstance(saved, list) else saved["files"]
        sources = [] if isinstance(saved, list) else saved.get("source_details", [])
        if any('\\\\win-server\\' in row['file_path'] for row in results):
            raise ValueError("已完成输出不可作为运行断点")
    image_inputs = []

    def progress(percent, message):
        progress_file.write_text(json.dumps({"task_id": task_id, "display_no": display_no,
            "progress": percent, "message": message, "finished_files": len(results),
            "total_files": len(candidates)}, ensure_ascii=False), encoding="utf-8")
        with SessionLocal() as db:
            task_repo.update_task_progress(db, task_id, progress=percent, message=message, status="running")
        print(json.dumps({"progress": percent, "message": message}, ensure_ascii=True), flush=True)

    def checkpoint():
        (output / "文件结果检查点.json").write_text(json.dumps({"files":results,"source_details":sources}, ensure_ascii=False), encoding="utf-8")

    def count_native(path):
        return service._count_single_file(file_path=path, root=root, converted_dir=converted,
            max_bytes=settings.WORD_COUNT_MAX_FILE_MB * 1024 * 1024, ocr_enabled=False)

    def reuse_pdf(path, row):
        key = normalized(original_path(path))
        detail = source_pages.get(key, [])
        if len(detail) != int(row["页数"]) or sorted(int(p["页码"]) for p in detail) != list(range(1, int(row["页数"])+1)):
            raise ValueError("PDF逐页数据不完整：" + str(path))
        prefix = "合计候选_"
        for field in ["word_count", "latin_word_count", "number_token_count", "mixed_latin_number_count", "han_count", "billable_latin_count", "billable_chinese_count"]:
            if sum(int(p[prefix+field]) for p in detail) != int(row[prefix+field]):
                raise ValueError("PDF逐页合计不一致：" + str(path))
        base = service._base_file_result(path, str(path.relative_to(root)), ".pdf", service.STATUS_COUNTED,
            size_bytes=path.stat().st_size, modified_at=datetime.fromtimestamp(path.stat().st_mtime).isoformat())
        buckets = dict(base["script_counts"])
        for field in ["latin_word_count", "number_token_count", "mixed_latin_number_count", "han_count"]:
            buckets[field] = int(row[prefix+field])
        buckets["cjk_punct_count"] = int(row[prefix+"billable_chinese_count"]) - buckets["han_count"]
        total = int(row[prefix+"word_count"])
        residual = total - sum(buckets.values())
        if residual < 0:
            raise ValueError("导入的脚本计数大于总数")
        buckets["other_count"] = residual
        quotes = service._quote_counts_from_script_counts(buckets)
        warning = "复用GPU报表的候选数量，未经最终人工验收；未合并内容未追加。" + row["复核原因"]
        if residual:
            warning += f" 原表缺少其他文字体系细分，{residual}项暂列其他。"
        base.update(word_count=total, main_word_count=total, page_count=int(row["页数"]),
            script_counts=buckets, script_count_total=total, quote_counts=quotes,
            ocr_used=any(p["文本来源"] != "文字层" for p in detail),
            ocr_page_count=sum(p["文本来源"] != "文字层" for p in detail),
            ocr_model="PaddleOCR GPU结果复用", ocr_review_pages=[int(p["页码"]) for p in detail if p["处理状态"] in ("review", "no_text_unverified")],
            stat_method="复用PDF GPU统计（文字层+OCR补充）", warning=warning, message="复用已有PDF候选统计",
            pdf_reuse_source_id=row["本地编号"], pdf_sha256=row["SHA256"], pdf_reuse_original=row,
            source_counts={"pdf_gpu_report_reuse": len(detail)}, **buckets, **quotes)
        # 原表未提供字符数、行数，保持空白，避免显示为真正的零。
        for field in ["char_count_no_spaces", "char_count_with_spaces", "cjk_char_count_no_spaces", "cjk_char_count_with_spaces",
                      "latin_char_count_no_spaces", "latin_char_count_with_spaces", "non_space_chars", "raw_chars", "line_count", "paragraph_count", "image_count"]:
            base[field] = None
        return base

    try:
        progress(2, f"发现{len(candidates)}个文件；正在复用PDF和解析Office/CAD")
        for index, path in enumerate(candidates, 1):
            extension = path.suffix.lower()
            if extension in service.IMAGE_EXTENSIONS:
                image_inputs.append({"index": index, "path": str(path)})
                continue
            if any(row.get("inventory_index") == index for row in results):
                continue
            row = lookup.get(normalized(original_path(path))) if extension == ".pdf" else None
            if row and sha(path) == row["SHA256"]:
                result = reuse_pdf(path, row)
                result["inventory_index"] = index
                results.append(result)
            else:
                result, rows, _ = count_native(path)
                result["inventory_index"] = index
                if extension == ".pdf":
                    result["warning"] = "当前PDF与GPU结果版本不一致或无匹配记录；仅提取当前文字层，未复用OCR补充。" + result.get("warning", "")
                    result["pdf_reuse_source_id"] = row["本地编号"] if row else ""
                    result["pdf_reuse_version_mismatch"] = bool(row)
                results.append(result)
                sources.extend(rows)
            checkpoint()
            progress(min(20, 2+len(results)*18//max(1,len(candidates)-len(image_inputs))), f"已统计：{path.name}")

        request = output / "image_request.json"
        request.write_text(json.dumps({"output_dir": str(output.resolve()), "images": image_inputs}, ensure_ascii=False), encoding="utf-8")
        worker = Path(__file__).with_name("arq_image_worker.py")
        with (output / "图片OCR运行日志.txt").open("w", encoding="utf-8") as log:
            process = subprocess.Popen([settings.PADDLEOCR_PYTHON, str(worker), str(request.resolve())],
                cwd=Path(__file__).resolve().parents[1], stdout=subprocess.PIPE, stderr=log, text=True, encoding="utf-8")
            for line in process.stdout:
                try:
                    event = json.loads(line)
                except ValueError:
                    log.write(line)
                    continue
                if "done" in event:
                    progress(20+event["done"]*65//max(1,event["total"]), f"图片OCR {event['done']}/{event['total']}：{Path(event['path']).name}")
            if process.wait():
                raise RuntimeError("图片识别进程失败，请检查运行日志")

        # 直接复用刚保存的OCR文本，调用原系统计数与导出逻辑。
        original_extract = service.extract_paddle_plain_text
        image_by_path = {normalized(row["path"]): row for row in image_inputs}

        def cached_image(**kwargs):
            image = image_by_path[normalized(kwargs["file_path"])]
            return json.loads((output / "image_results" / f"{image['index']:03d}.json").read_text(encoding="utf-8"))

        service.extract_paddle_plain_text = cached_image
        for index, image in enumerate(image_inputs, 1):
            path = Path(image["path"])
            result, rows, _ = service._count_single_file(file_path=path, root=root, converted_dir=converted,
                max_bytes=settings.WORD_COUNT_MAX_FILE_MB*1024*1024, ocr_enabled=True, ocr_model="local/paddleocr",
                ocr_text_dir=output / "OCR识别文本")
            result["inventory_index"] = image["index"]
            results.append(result)
            sources.extend(rows)
            checkpoint()
        service.extract_paddle_plain_text = original_extract
        results.sort(key=lambda row: row["inventory_index"])
        for row in results + sources:
            if row.get("relative_path"):
                row["file_path"] = str(display_root / row["relative_path"])
        summary = service._build_summary(results, truncated=False, started_at=started)
        summary["character_totals_incomplete"] = True
        summary["pdf_reused_files"] = sum("pdf_reuse_original" in row for row in results)
        summary["pdf_version_mismatch_files"] = sum(bool(row.get("pdf_reuse_version_mismatch")) for row in results)
        summary["extension_counts"] = dict(Counter(row["extension"] for row in results))
        report = {"task_id": task_id, "directory_path": str(display_root), "input_path": str(display_root), "input_kind": "directory",
            "input_source": "path", "summary": summary, "files": results, "source_details": sources,
            "generated_at": datetime.now().isoformat(), "rules": [
                "PDF按共享路径和SHA256匹配复用GPU候选统计，未重新OCR或重复叠加文字层。",
                "版本不一致PDF仅提取当前文字层，OCR补充未复用。",
                "图片重新使用PaddleOCR识别，按原图分辨率处理，最高300DPI，低于0.6置信度的识别行不计入。",
                "DWG通过ODA转DXF，统计模型空间、图纸空间和插入块文字；不自动读取外部参照。",
                "复用PDF没有字符数、行数等原始明细，相关列空白，字符数汇总不完整，不用于报价。",
                "候选报价量保留重复文件和PDF/DWG对应内容，未按交付版本去重。",
                "OCR复核提示、编码异常、版本问题和未合并内容仍需人工核验。"]}
        if previous:
            report["rules"].append("断点续跑保留已完成的逐文件统计；非图片来源细分见文件明细。")
        report["report_excel"] = service._output_web_path(output / "字数统计报告.xlsx")
        report["report_json"] = service._output_web_path(output / "字数统计结果.json")
        report["summary_text"] = f"统计{len(results)}份：拉丁系候选{summary['total_billable_latin_count']}，PDF复用{summary['pdf_reused_files']}份"
        service._write_excel_report(output / "字数统计报告.xlsx", report)
        (output / "字数统计结果.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        review = [row for row in results if row.get("warning") or row["status"] != "counted"]
        with (output / "复核清单.csv").open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=["file_path", "relative_path", "status", "word_count", "billable_latin_count", "warning", "error"], extrasaction="ignore")
            writer.writeheader()
            writer.writerows(review)
        outputs = [{"name": "字数统计报告.xlsx", "path": report["report_excel"]},
                   {"name": "字数统计结果.json", "path": report["report_json"]},
                   {"name": "复核清单.csv", "path": service._output_web_path(output / "复核清单.csv")}]
        with SessionLocal() as db:
            task_repo.complete_task(db, task_id, result_json=json.dumps(report, ensure_ascii=False),
                output_path=report["report_excel"], output_files_json=json.dumps(outputs, ensure_ascii=False), message=report["summary_text"])
        progress_file.write_text(json.dumps({"task_id": task_id, "display_no": display_no, "progress": 100, "summary": summary,
            "output_dir": str(output.resolve())}, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"done": True, "task_id": task_id, "display_no": display_no, "summary": summary}, ensure_ascii=True), flush=True)
    except Exception as exc:
        with SessionLocal() as db:
            task_repo.fail_task(db, task_id, str(exc))
        raise


if __name__ == "__main__":
    main()
