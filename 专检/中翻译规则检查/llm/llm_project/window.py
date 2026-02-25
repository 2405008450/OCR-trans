import os
import sys
import json
import time
import subprocess
import threading
import traceback
import logging
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document

from llm.llm_project.llm_check.check import Match
from llm.llm_project.parsers.body_extractor import extract_body_text
from llm.llm_project.parsers.footer_extractor import extract_footers
from llm.llm_project.parsers.header_extractor import extract_headers
from llm.llm_project.replace.fix_replace_json import ensure_backup_copy, CommentManager, replace_and_comment_in_docx
from llm.zhongfanyi.clean_json import load_json_file
from llm.zhongfanyi.json_files import write_json_with_timestamp

APP_NAME = "译文审校与自动替换工具（右侧文件面板）"
DEFAULT_CONFIG_PATH = "config.json"


# ========= 通用工具 =========
def open_path(path: str):
    if not path or not os.path.exists(path):
        raise FileNotFoundError(f"路径不存在：{path}")
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


def open_folder(path: str):
    if not path:
        raise FileNotFoundError("路径为空")
    if not os.path.exists(path):
        raise FileNotFoundError(f"路径不存在：{path}")
    if os.path.isfile(path):
        path = os.path.dirname(path)
    open_path(path)


def sizeof_fmt(num_bytes: int) -> str:
    if num_bytes < 0:
        return "0 B"
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(num_bytes)
    for u in units:
        if size < 1024.0 or u == units[-1]:
            return f"{int(size)} B" if u == "B" else f"{size:.1f} {u}"
        size /= 1024.0
    return f"{int(num_bytes)} B"


def safe_filename(name: str) -> str:
    # 简单清理非法字符（Windows 为主），避免生成文件失败
    bad = '<>:"/\\|?*'
    for c in bad:
        name = name.replace(c, "_")
    return name.strip() or "output"


def load_config(path=DEFAULT_CONFIG_PATH) -> dict:
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(cfg: dict, path=DEFAULT_CONFIG_PATH):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def setup_logger(output_dir: str, log_filename: str):
    os.makedirs(output_dir, exist_ok=True)
    log_path = os.path.join(output_dir, log_filename)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler()],
    )
    return log_path


# ========= 右侧文件面板 =========
class FilePanel(ttk.Frame):
    """
    右侧常驻文件窗口：
    - 显示输出目录文件
    - 双击打开
    - 右键：打开/打开文件夹/复制路径/刷新/删除(可选)
    - 支持“只看本工具生成文件”过滤
    """

    def __init__(self, master, get_output_dir_callable, get_filter_prefixes_callable):
        super().__init__(master)
        self.get_output_dir = get_output_dir_callable
        self.get_filter_prefixes = get_filter_prefixes_callable

        self.var_only_tool_files = tk.BooleanVar(value=True)
        self._build_ui()
        self.refresh()

    def _build_ui(self):
        pad = 6

        header = ttk.Frame(self)
        header.pack(fill=tk.X, padx=pad, pady=(pad, 0))

        ttk.Label(header, text="输出文件", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)

        ttk.Button(header, text="刷新", width=6, command=self.refresh).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(header, text="打开目录", width=8, command=self.open_output_dir).pack(side=tk.RIGHT)

        opt = ttk.Frame(self)
        opt.pack(fill=tk.X, padx=pad, pady=(6, 0))
        ttk.Checkbutton(opt, text="仅显示本工具生成文件", variable=self.var_only_tool_files, command=self.refresh).pack(
            side=tk.LEFT
        )

        columns = ("name", "mtime", "size", "path")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=16)
        self.tree.heading("name", text="文件名")
        self.tree.heading("mtime", text="修改时间")
        self.tree.heading("size", text="大小")
        self.tree.heading("path", text="路径")
        self.tree.column("name", width=180, anchor="w")
        self.tree.column("mtime", width=145, anchor="w")
        self.tree.column("size", width=80, anchor="e")
        self.tree.column("path", width=260, anchor="w")

        yscroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(pad, 0), pady=pad)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y, pady=pad, padx=(0, pad))

        self.tree.bind("<Double-1>", lambda _e: self.open_selected_file())
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Button-2>", self._on_right_click)

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="打开文件", command=self.open_selected_file)
        self.menu.add_command(label="打开所在文件夹", command=self.open_selected_folder)
        self.menu.add_separator()
        self.menu.add_command(label="复制路径", command=self.copy_selected_path)
        self.menu.add_separator()
        self.menu.add_command(label="刷新", command=self.refresh)

        self.status = tk.StringVar(value="—")
        ttk.Label(self, textvariable=self.status, anchor="w").pack(fill=tk.X, padx=pad, pady=(0, pad))

    def _get_selected_path(self):
        sel = self.tree.selection()
        if not sel:
            return None
        values = self.tree.item(sel[0], "values")
        if not values:
            return None
        return values[3]

    def _match_filter(self, filename: str) -> bool:
        if not self.var_only_tool_files.get():
            return True
        prefixes = self.get_filter_prefixes() or []
        if not prefixes:
            return True
        return any(filename.startswith(p) for p in prefixes)

    def refresh(self):
        out_dir = (self.get_output_dir() or "").strip()
        self.tree.delete(*self.tree.get_children())

        if not out_dir or not os.path.isdir(out_dir):
            self.status.set("输出目录未设置或无效")
            return

        items = []
        try:
            for fn in os.listdir(out_dir):
                full = os.path.join(out_dir, fn)

                # 如果是文件夹：直接添加，不进行“工具生成文件”过滤（方便导航）
                if os.path.isdir(full):
                    st = os.stat(full)
                    items.append({
                        "type": "folder",
                        "name": f"📁 {fn}",  # 添加图标前缀
                        "full": full,
                        "mtime": st.st_mtime,
                        "size_str": "--"
                    })
                # 如果是文件：执行过滤逻辑
                elif os.path.isfile(full):
                    if not self._match_filter(fn):
                        continue
                    st = os.stat(full)
                    items.append({
                        "type": "file",
                        "name": f"📄 {fn}",
                        "full": full,
                        "mtime": st.st_mtime,
                        "size_str": sizeof_fmt(int(st.st_size))
                    })
        except Exception as e:
            self.status.set(f"读取目录失败：{e}")
            return

        # 排序逻辑：先按类型(文件夹在前)，再按时间(最新在前)
        items.sort(key=lambda x: (x["type"] != "folder", -x["mtime"]))

        for item in items:
            mtime_str = datetime.fromtimestamp(item["mtime"]).strftime("%Y-%m-%d %H:%M:%S")
            self.tree.insert("", "end", values=(item["name"], mtime_str, item["size_str"], item["full"]))

        self.status.set(f"{out_dir} | 总数：{len(items)}")

    def open_output_dir(self):
        out_dir = (self.get_output_dir() or "").strip()
        if not out_dir:
            messagebox.showwarning("未设置输出目录", "请先在左侧选择输出文件夹。")
            return
        try:
            open_folder(out_dir)
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    def open_selected_file(self):
        p = self._get_selected_path()
        if not p:
            return
        try:
            if os.path.isdir(p):
                # 如果选中的是文件夹，直接打开文件夹
                open_folder(p)
            else:
                # 如果是文件，调用系统关联程序打开
                open_path(p)
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    def open_selected_folder(self):
        p = self._get_selected_path()
        if not p:
            return
        try:
            open_folder(p)
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    def copy_selected_path(self):
        p = self._get_selected_path()
        if not p:
            return
        self.clipboard_clear()
        self.clipboard_append(p)
        self.status.set(f"已复制路径：{p}")

    def _on_right_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()


# ========= 主应用 =========
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1120x560")
        self.resizable(True, True)

        self.cfg = load_config()
        self._build_ui()
        self._load_cfg_to_ui()

    def _build_ui(self):
        pad = 10

        root = ttk.Frame(self, padding=pad)
        root.pack(fill=tk.BOTH, expand=True)

        # 左右分栏
        paned = ttk.Panedwindow(root, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # ===== 左侧：配置 & 日志 =====
        left = ttk.Frame(paned)
        paned.add(left, weight=3)

        # 文件选择
        file_box = ttk.LabelFrame(left, text="文件与输出", padding=pad)
        file_box.pack(fill=tk.X)

        self.var_src = tk.StringVar()
        self.var_tgt = tk.StringVar()
        self.var_out = tk.StringVar()

        self._row_file(file_box, "原文（docx）", self.var_src, self._pick_src, 0)
        self._row_file(file_box, "译文（docx）", self.var_tgt, self._pick_tgt, 1)
        self._row_dir(file_box, "输出文件夹", self.var_out, self._pick_out, 2)

        # 文件名模板设置（你说的“可设置文本对比结果等所有生成文件”）
        name_box = ttk.LabelFrame(left, text="生成文件命名（可改）", padding=pad)
        name_box.pack(fill=tk.X, pady=(pad, 0))

        self.var_report_name = tk.StringVar()   # 文本对比结果
        self.var_fixed_name = tk.StringVar()    # 修订译文
        self.var_log_name = tk.StringVar()      # 日志

        ttk.Label(name_box, text="对比结果").grid(row=0, column=0, sticky="w")
        ttk.Entry(name_box, textvariable=self.var_report_name, width=64).grid(row=0, column=1, padx=8, pady=4, sticky="w")
        ttk.Label(name_box, text="支持 {ts} 时间戳").grid(row=0, column=2, sticky="w")

        ttk.Label(name_box, text="修订译文").grid(row=1, column=0, sticky="w")
        ttk.Entry(name_box, textvariable=self.var_fixed_name, width=64).grid(row=1, column=1, padx=8, pady=4, sticky="w")
        ttk.Label(name_box, text="例如 译文_自动修订_{ts}.docx").grid(row=1, column=2, sticky="w")

        ttk.Label(name_box, text="日志文件").grid(row=2, column=0, sticky="w")
        ttk.Entry(name_box, textvariable=self.var_log_name, width=64).grid(row=2, column=1, padx=8, pady=4, sticky="w")
        ttk.Label(name_box, text="例如 app.log").grid(row=2, column=2, sticky="w")

        # 模型配置区
        model_box = ttk.LabelFrame(left, text="大模型配置", padding=pad)
        model_box.pack(fill=tk.X, pady=(pad, 0))

        self.var_api_key = tk.StringVar()
        self.var_base_url = tk.StringVar()
        self.var_model = tk.StringVar()

        self._row_entry(model_box, "API Key", self.var_api_key, 0, show="*")
        self._row_entry(model_box, "Base URL", self.var_base_url, 1)
        self._row_entry(model_box, "Model", self.var_model, 2)

        # 操作区
        action_box = ttk.Frame(left)
        action_box.pack(fill=tk.X, pady=(pad, 0))

        self.btn_run = ttk.Button(action_box, text="开始：AI审校 → 解析 → 自动替换", command=self.on_run)
        self.btn_run.pack(side=tk.LEFT)

        self.btn_save = ttk.Button(action_box, text="保存配置", command=self.on_save_cfg)
        self.btn_save.pack(side=tk.LEFT, padx=(pad, 0))

        # 日志区
        log_box = ttk.LabelFrame(left, text="运行日志", padding=pad)
        log_box.pack(fill=tk.BOTH, expand=True, pady=(pad, 0))

        self.txt_log = tk.Text(log_box, height=10, wrap="word")
        self.txt_log.pack(fill=tk.BOTH, expand=True)

        self._log("就绪。左侧配置；右侧为输出文件窗口（会自动刷新）。")

        # ===== 右侧：常驻文件面板 =====
        right = ttk.Frame(paned)
        paned.add(right, weight=2)

        self.file_panel = FilePanel(
            right,
            get_output_dir_callable=self._get_output_dir,
            get_filter_prefixes_callable=self._get_tool_file_prefixes,
        )
        self.file_panel.pack(fill=tk.BOTH, expand=True)

    def _row_file(self, parent, label, var, cmd, r):
        ttk.Label(parent, text=label, width=12).grid(row=r, column=0, sticky="w")
        ttk.Entry(parent, textvariable=var, width=70).grid(row=r, column=1, padx=8, pady=4, sticky="w")
        ttk.Button(parent, text="选择", command=cmd, width=8).grid(row=r, column=2, sticky="e")

    def _row_dir(self, parent, label, var, cmd, r):
        ttk.Label(parent, text=label, width=12).grid(row=r, column=0, sticky="w")
        ttk.Entry(parent, textvariable=var, width=70).grid(row=r, column=1, padx=8, pady=4, sticky="w")
        ttk.Button(parent, text="选择", command=cmd, width=8).grid(row=r, column=2, sticky="e")

    def _row_entry(self, parent, label, var, r, show=None):
        ttk.Label(parent, text=label, width=12).grid(row=r, column=0, sticky="w")
        ent = ttk.Entry(parent, textvariable=var, width=70, show=show)
        ent.grid(row=r, column=1, padx=8, pady=4, sticky="w")

    def _pick_src(self):
        p = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        if p:
            self.var_src.set(p)

    def _pick_tgt(self):
        p = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        if p:
            self.var_tgt.set(p)

    def _pick_out(self):
        p = filedialog.askdirectory()
        if p:
            self.var_out.set(p)
            self.file_panel.refresh()

    def _get_output_dir(self) -> str:
        return self.var_out.get()

    def _get_tool_file_prefixes(self):
        # 右侧“仅显示本工具生成文件”过滤用：取你当前配置的文件名前缀
        # 例：文本对比结果_{ts}.docx -> 前缀“文本对比结果_”
        def prefix_from_template(tpl: str) -> str:
            tpl = (tpl or "").strip()
            if "{ts}" in tpl:
                return tpl.split("{ts}")[0]
            # 没写 {ts} 的情况下，取去掉扩展名的部分作为前缀
            return os.path.splitext(tpl)[0]

        return [
            prefix_from_template(self.var_report_name.get()),
            prefix_from_template(self.var_fixed_name.get()),
            os.path.splitext(self.var_log_name.get().strip() or "app.log")[0],
            "zhengwen",  # 文件夹名
            "yemei",  # 文件夹名
            "yejiao"  # 文件夹名
        ]

    def _load_cfg_to_ui(self):
        self.var_api_key.set(self.cfg.get("api_key", ""))
        self.var_base_url.set(self.cfg.get("base_url", ""))
        self.var_model.set(self.cfg.get("model", "google/gemini-2.5-pro"))
        self.var_out.set(self.cfg.get("output_dir", ""))

        # 默认命名模板
        self.var_report_name.set(self.cfg.get("report_name", "文本对比结果_{ts}.docx"))
        self.var_fixed_name.set(self.cfg.get("fixed_name", "译文_自动修订_{ts}.docx"))
        self.var_log_name.set(self.cfg.get("log_name", "app.log"))

        # 初始化刷新右侧文件面板
        self.after(200, self.file_panel.refresh)

    def on_save_cfg(self):
        self.cfg["api_key"] = self.var_api_key.get().strip()
        self.cfg["base_url"] = self.var_base_url.get().strip()
        self.cfg["model"] = self.var_model.get().strip()
        self.cfg["output_dir"] = self.var_out.get().strip()

        self.cfg["report_name"] = self.var_report_name.get().strip()
        self.cfg["fixed_name"] = self.var_fixed_name.get().strip()
        self.cfg["log_name"] = self.var_log_name.get().strip()

        save_config(self.cfg)
        messagebox.showinfo("已保存", "配置已保存到 config.json")
        self.file_panel.refresh()

    def on_run(self):
        src = self.var_src.get().strip()
        tgt = self.var_tgt.get().strip()
        out = self.var_out.get().strip()

        api_key = self.var_api_key.get().strip()
        base_url = self.var_base_url.get().strip()
        model = self.var_model.get().strip()

        if not (src and tgt and out):
            messagebox.showwarning("缺少参数", "请先选择原文、译文和输出文件夹。")
            return
        if not os.path.exists(src) or not os.path.exists(tgt):
            messagebox.showerror("文件不存在", "原文或译文路径无效。")
            return
        if not api_key or not base_url or not model:
            messagebox.showwarning("缺少模型配置", "请填写 API Key / Base URL / Model。")
            return

        self.btn_run.config(state=tk.DISABLED)
        self._log("开始运行...")
        self.file_panel.refresh()

        t = threading.Thread(
            target=self._run_pipeline_safe,
            args=(src, tgt, out, api_key, base_url, model),
            daemon=True,
        )
        t.start()

    def _render_name(self, template: str, ts: str) -> str:
        template = safe_filename(template.strip())
        # 允许用户不写扩展名：对比结果/修订译文默认 docx
        return template.replace("{ts}", ts)

    def _run_pipeline_safe(self, src, tgt, out, api_key, base_url, model):
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")

            report_name = self._render_name(self.var_report_name.get() or "文本对比结果_{ts}.docx", ts)
            fixed_name = self._render_name(self.var_fixed_name.get() or "译文_自动修订_{ts}.docx", ts)
            log_name = safe_filename(self.var_log_name.get().strip() or "app.log")

            # 兜底扩展名
            if not report_name.lower().endswith(".docx"):
                report_name += ".docx"
            if not fixed_name.lower().endswith(".docx"):
                fixed_name += ".docx"

            report_docx = os.path.join(out, report_name)
            fixed_docx = os.path.join(out, fixed_name)

            log_path = setup_logger(out, log_name)
            self._log(f"日志文件：{log_path}")

            # ===== 1) AI审校生成报告（对接你的 matcher_1.py）=====
            self._log("步骤1/3：AI审校生成错误报告...")
            original_path = src  # 请替换为原文文件路径
            translated_path = tgt  # 请替换为译文文件路径
            # 处理页眉
            original_header_text = extract_headers(original_path)
            translated_header_text = extract_headers(translated_path)
            # 处理页脚
            original_footer_text = extract_footers(original_path)
            translated_footer_text = extract_footers(translated_path)
            # 处理正文(含脚注/表格/自动编号)
            original_body_text = extract_body_text(original_path)
            translated_body_text = extract_body_text(translated_path)

            # 1. 定义正文、页眉、页脚的输出子目录
            body_out_dir = os.path.join(out, "zhengwen", "output_json")
            header_out_dir = os.path.join(out, "yemei", "output_json")
            footer_out_dir = os.path.join(out, "yejiao", "output_json")

            # 2. 确保这些文件夹存在（如果不存在则自动创建）
            os.makedirs(body_out_dir, exist_ok=True)
            os.makedirs(header_out_dir, exist_ok=True)
            os.makedirs(footer_out_dir, exist_ok=True)

            # 实例化对象并进行对比
            matcher = Match()
            # 正文对比
            print("======正在检查正文===========")
            if original_body_text and translated_body_text:
                # 两个值都不为空，正常执行比较
                body_result = matcher.compare_texts(original_body_text, translated_body_text)
            else:
                # 任意一个为空，生成空结果
                body_result = {}  # 或者 body_result = []，根据你的 write_json_with_timestamp 函数期望的格式
                print("原文或译文为空，检查结果为空")
            self._log("正在生成正文中间 JSON 报告...")
            body_result_name, body_result_path = write_json_with_timestamp(
                body_result,
                body_out_dir
            )
            # body_result = matcher.compare_texts(original_body_text, translated_body_text)
            # body_result_name, body_result_path = write_json_with_timestamp(body_result,r"C:\Users\Administrator\Desktop\project\llm\llm_project\zhengwen\output_json")
            # #页眉对比
            print("======正在检查页眉===========")
            if original_header_text and translated_header_text:
                # 两个值都不为空，正常执行比较
                header_result = matcher.compare_texts(original_header_text, translated_header_text)
            else:
                # 任意一个为空，生成空结果
                header_result = {}
                print("原文或译文为空，检查结果为空")
            self._log("正在生成页眉中间 JSON 报告...")
            header_result_name, header_result_path = write_json_with_timestamp(
                header_result,
                header_out_dir
            )
            # header_result = matcher.compare_texts(original_header_text, translated_header_text)
            # header_result_name, header_result_path = write_json_with_timestamp(header_result, r"C:\Users\Administrator\Desktop\project\llm\llm_project\yemei\output_json")
            # #页脚对比
            print("======正在检查页脚===========")
            if original_footer_text and translated_footer_text:
                # 两个值都不为空，正常执行比较
                footer_result = matcher.compare_texts(original_footer_text, translated_footer_text)
            else:
                # 任意一个为空，生成空结果
                footer_result = {}
                print("原文或译文为空，检查结果为空")
            self._log("正在生成页脚中间 JSON 报告...")
            footer_result_name, footer_result_path = write_json_with_timestamp(
                footer_result,
                footer_out_dir
            )
            # footer_result = matcher.compare_texts(original_footer_text, translated_footer_text)
            # footer_result_name, footer_result_path = write_json_with_timestamp(footer_result, r"C:\Users\Administrator\Desktop\project\llm\llm_project\yejiao\output_json")

            print("================================")
            print("请把 run_ai_check(...) 对接到 matcher_1.py 的实际函数")

            # ===== 2) 解析报告 + 自动替换（对接你的 1.py）=====
            # 1) 复制译文到 backup/
            backup_copy_path = ensure_backup_copy(translated_path)
            print(f"✅ 已复制译文副本到: {backup_copy_path}")

            # 2) 读取错误报告并解析
            print("\n正在提取解析正文错误报告...")
            body_errors = load_json_file(body_result_path)
            print("正文错误报告", body_errors)
            for err in body_errors:
                print(err)
            print("正文错误解析个数：", len(body_errors))

            print("\n正在提取解析页眉错误报告...")
            header_errors = load_json_file(header_result_path)
            print("页眉错误报告", header_errors)
            print("页眉错误解析个数：", len(header_errors))

            print("\n正在提取解析页脚错误报告...")
            footer_errors = load_json_file(footer_result_path)
            print("页脚错误报告", footer_errors)
            print("页脚错误解析个数：", len(footer_errors))

            # 3) 打开副本 docx
            print("正在加载文档...")
            doc = Document(backup_copy_path)

            # 4) 创建批注管理器并初始化
            print("正在初始化批注系统...")
            comment_manager = CommentManager(doc)

            # 【关键】创建初始批注以确保 comments.xml 结构完整
            if comment_manager.create_initial_comment():
                print("✓ 批注系统初始化成功\n")
            else:
                print("⚠️ 批注系统初始化失败，但将继续尝试处理\n")

            # 5) 逐条执行替换并添加批注
            print("==================== 开始处理正文错误 ====================\n")
            body_success_count = 0
            body_fail_count = 0

            for idx, e in enumerate(body_errors, 1):
                err_id = e.get("错误编号", "?")
                err_type = e.get("错误类型", "")
                old = (e.get("译文数值") or "").strip()
                new = (e.get("译文修改建议值") or "").strip()
                reason = (e.get("修改理由"), "")
                trans_context = e.get("译文上下文", "") or ""
                anchor = e.get("替换锚点", "") or ""

                if not old or not new:
                    print(f"[{idx}/{len(body_errors)}] [跳过] 错误 #{err_id}: old/new 缺失")
                    body_fail_count += 1
                    continue

                # 执行替换并添加批注(正文)
                ok, strategy = replace_and_comment_in_docx(
                    doc, old, new, reason, comment_manager,
                    context=trans_context,
                    anchor_text=anchor
                )

                if ok:
                    print(f"[{idx}/{len(body_errors)}] [✓成功] 错误 #{err_id} ({err_type})")
                    print(f"    策略: {strategy}")
                    print(f"    修改理由: {reason}")
                    print(f"    操作: '{old}' → '{new}' (已替换并添加批注)")
                    if anchor:
                        print(f"    锚点: {anchor}...")
                    elif trans_context:
                        print(f"    上下文: {trans_context}...")
                    body_success_count += 1
                else:
                    print(f"[{idx}/{len(body_errors)}] [✗失败] 错误 #{err_id} ({err_type})")
                    print(f"    未找到匹配: '{old}'")
                    if anchor:
                        print(f"    锚点: {anchor}...")
                    print(f"    上下文: {trans_context if trans_context else '无'}...")
                    body_fail_count += 1
                print()

            print("==================== 开始处理页眉错误 ====================\n")
            header_success_count = 0
            header_fail_count = 0

            for idx, e in enumerate(header_errors, 1):
                err_id = e.get("错误编号", "?")
                err_type = e.get("错误类型", "")
                old = (e.get("译文数值") or "").strip()
                new = (e.get("译文修改建议值") or "").strip()
                reason = (e.get("修改理由"), "")
                trans_context = e.get("译文上下文", "") or ""
                anchor = e.get("替换锚点", "") or ""

                if not old or not new:
                    print(f"[{idx}/{len(header_errors)}] [跳过] 错误 #{err_id}: old/new 缺失")
                    header_fail_count += 1
                    continue

                # 执行替换并添加批注(正文)
                ok, strategy = replace_and_comment_in_docx(
                    doc, old, new, reason, comment_manager,
                    context=trans_context,
                    anchor_text=anchor
                )

                if ok:
                    print(f"[{idx}/{len(header_errors)}] [✓成功] 错误 #{err_id} ({err_type})")
                    print(f"    修改理由: {reason}")
                    print(f"    策略: {strategy}")
                    print(f"    操作: '{old}' → '{new}' (已替换并添加批注)")
                    if anchor:
                        print(f"    锚点: {anchor}...")
                    elif trans_context:
                        print(f"    上下文: {trans_context}...")
                    header_success_count += 1
                else:
                    print(f"[{idx}/{len(header_errors)}] [✗失败] 错误 #{err_id} ({err_type})")
                    print(f"    未找到匹配: '{old}'")
                    if anchor:
                        print(f"    锚点: {anchor}...")
                    print(f"    上下文: {trans_context if trans_context else '无'}...")
                    header_fail_count += 1
                print()

            print("==================== 开始处理页脚错误 ====================\n")
            footer_success_count = 0
            footer_fail_count = 0

            for idx, e in enumerate(footer_errors, 1):
                err_id = e.get("错误编号", "?")
                err_type = e.get("错误类型", "")
                old = (e.get("译文数值") or "").strip()
                new = (e.get("译文修改建议值") or "").strip()
                reason = (e.get("修改理由"), "")
                trans_context = e.get("译文上下文", "") or ""
                anchor = e.get("替换锚点", "") or ""

                if not old or not new:
                    print(f"[{idx}/{len(footer_errors)}] [跳过] 错误 #{err_id}: old/new 缺失")
                    footer_fail_count += 1
                    continue

                # 执行替换并添加批注(正文)
                ok, strategy = replace_and_comment_in_docx(
                    doc, old, new, reason, comment_manager,
                    context=trans_context,
                    anchor_text=anchor
                )

                if ok:
                    print(f"[{idx}/{len(footer_errors)}] [✓成功] 错误 #{err_id} ({err_type})")
                    print(f"    修改理由: {reason}")
                    print(f"    策略: {strategy}")
                    print(f"    操作: '{old}' → '{new}' (已替换并添加批注)")
                    if anchor:
                        print(f"    锚点: {anchor}...")
                    elif trans_context:
                        print(f"    上下文: {trans_context}...")
                    footer_success_count += 1
                else:
                    print(f"[{idx}/{len(footer_errors)}] [✗失败] 错误 #{err_id} ({err_type})")
                    print(f"    未找到匹配: '{old}'")
                    if anchor:
                        print(f"    锚点: {anchor}...")
                    print(f"    上下文: {trans_context if trans_context else '无'}...")
                    footer_fail_count += 1
                print()

            # 6) 保存文档
            print("正在保存文档...")
            doc.save(backup_copy_path)

            print(f"\n==================== 正文处理完成 ====================")
            print(f"成功: {body_success_count} | 失败: {body_fail_count} | 总计: {len(body_errors)}")
            if len(body_errors) > 0:
                print(f"成功率: {body_success_count / len(body_errors) * 100:.1f}%")
            print(f"\n✅ 已保存到: {backup_copy_path}")
            print("⚠️ 原始译文文件未被修改")

            print(f"\n==================== 页眉处理完成 ====================")
            print(f"成功: {header_success_count} | 失败: {header_fail_count} | 总计: {len(header_errors)}")
            if len(header_errors) > 0:
                print(f"成功率: {header_success_count / len(header_errors) * 100:.1f}%")
            print(f"\n✅ 已保存到: {backup_copy_path}")
            print("⚠️ 原始译文文件未被修改")

            print(f"\n==================== 页脚处理完成 ====================")
            print(f"成功: {footer_success_count} | 失败: {footer_fail_count} | 总计: {len(footer_errors)}")
            if len(footer_errors) > 0:
                print(f"成功率: {footer_success_count / len(footer_errors) * 100:.1f}%")
            print(f"\n✅ 已保存到: {backup_copy_path}")
            print("⚠️ 原始译文文件未被修改")

            print(f"\n==================== 文章处理完成 ====================")
            count = len(body_errors) + len(header_errors) + len(footer_errors)
            success_count = body_success_count + header_success_count + footer_success_count
            fail_count = body_fail_count + header_fail_count + footer_fail_count
            print(f"成功: {success_count} | 失败: {fail_count} | 总计: {count}")
            if len(footer_errors) > 0:
                print(f"成功率: {success_count / count * 100:.1f}%")
            print(f"\n✅ 已保存到: {backup_copy_path}")
            print("⚠️ 原始译文文件未被修改")

            self._log(f"输出：{report_docx}")
            self._log(f"输出：{fixed_docx}")
            self._log("步骤3/3：完成 ✅")

            def _finish_ui():
                self.file_panel.refresh()
                messagebox.showinfo("完成", f"已完成。\n对比结果：{report_docx}\n修订译文：{fixed_docx}")

            self.after(0, _finish_ui)

        except Exception as e:
            tb = traceback.format_exc()
            logging.exception("运行失败")

            def _fail_ui():
                self._log("运行失败 ❌")
                self._log(str(e))
                self._log(tb)
                self.file_panel.refresh()
                messagebox.showerror("运行失败", f"{e}\n\n详细错误见日志/窗口。")

            self.after(0, _fail_ui)

        finally:

            def _enable_ui():
                self.btn_run.config(state=tk.NORMAL)
                self.file_panel.refresh()

            self.after(0, _enable_ui)

    def _log(self, msg: str):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")


if __name__ == "__main__":
    App().mainloop()
