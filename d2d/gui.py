# -*- coding: utf-8 -*-
"""Tkinter 图形界面。

两个页签:
  批量转换 —— 选择分区/文件夹/单文件 → 扫描 → 进度转换 → 导出日志
  测试校验 —— 生成测试 .docm → 转换 → 机器校验 → 人工复核
"""
from __future__ import annotations

import os
import queue
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from typing import Dict, List, Optional, Tuple

from . import scanner, search_kernel
from .constants import (
    CLEANUP_BY_LABEL, CLEANUP_DELETE, CLEANUP_LABELS, CLEANUP_NONE,
    CLEANUP_RECYCLE, OVERWRITE_BY_LABEL, OVERWRITE_LABELS, OVERWRITE_RENAME,
    OVERWRITE_REPLACE, OVERWRITE_SKIP,
)
from .search_kernel import (
    KERNEL_AUTO, KERNEL_BY_LABEL, KERNEL_EVERYTHING, KERNEL_LABELS,
    KERNEL_WALK, KERNEL_WINDOWS_INDEX,
)
from .engine import (
    ConvertItem, STATUS_FAILED, STATUS_PENDING, STATUS_SKIPPED, STATUS_SUCCESS,
    WordSession, convert_batch, plan_output,
)
from .testgen import create_test_docm
from .verify import verify_conversion

KIND_LABELS = {"drive": "分区", "folder": "文件夹", "file": "文件"}
STATUS_COLORS = {
    STATUS_PENDING: "#444444",
    STATUS_SUCCESS: "#137333",
    STATUS_SKIPPED: "#8a6d00",
    STATUS_FAILED: "#c5221f",
}


class ConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DOCM → DOCX 批量无损转换器")
        self.geometry("1240x880")
        self.minsize(1040, 720)

        self.queue: "queue.Queue[Tuple[str, object]]" = queue.Queue()
        self.cancel_event = threading.Event()
        self.busy = False
        self._poll_id = None

        self.targets: List[Dict[str, str]] = []      # {"kind","path"}
        self.files: List[Dict[str, str]] = []        # {"path","size","mtime","status","message","root"}
        self.file_records: Dict[str, Dict[str, str]] = {}   # path → 记录(实时统计用)
        self.row_ids: Dict[str, str] = {}            # path → tree iid
        self.counters = {STATUS_SUCCESS: 0, STATUS_SKIPPED: 0, STATUS_FAILED: 0}
        self.log_lines: List[str] = []

        self.test_dir = os.path.join(os.path.expanduser("~"), "docm2docx_test")
        self.test_docm: Optional[str] = None
        self.test_docx: Optional[str] = None

        self.var_drive = tk.StringVar()
        self.var_recursive = tk.BooleanVar(value=True)
        self.var_skip_system = tk.BooleanVar(value=True)
        self.var_unblock = tk.BooleanVar(value=True)
        self.var_out_same = tk.BooleanVar(value=True)
        self.var_out_dir = tk.StringVar()
        self.var_overwrite = tk.StringVar(value=OVERWRITE_LABELS[OVERWRITE_SKIP])
        self.var_cleanup = tk.StringVar(value=CLEANUP_LABELS[CLEANUP_NONE])
        self.var_kernel = tk.StringVar(value=KERNEL_LABELS[KERNEL_AUTO])
        self.var_status = tk.StringVar(value="就绪")
        self.var_scan_info = tk.StringVar(value="尚未扫描")
        self.var_counters = tk.StringVar(value="成功 0　跳过 0　失败 0　总计 0")
        self.var_speed = tk.StringVar(value="用时 00:00　速度 -")
        self.var_test_status = tk.StringVar(value="尚未生成测试文档")

        self._build_ui()
        self._refresh_drives()
        self.after(120, self._poll)

    # ===================================================================== #
    # 界面构建
    # ===================================================================== #
    def _build_ui(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Treeview", rowheight=22)
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 11, "bold"))

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)
        notebook.add(self._build_convert_tab(notebook), text="  批量转换  ")
        notebook.add(self._build_test_tab(notebook), text="  测试与校验  ")

    # ---------------------------------------------------------------- 转换页 #
    def _build_convert_tab(self, parent) -> ttk.Frame:
        page = ttk.Frame(parent, padding=8)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(3, weight=1)

        # ① 扫描目标 --------------------------------------------------- #
        box1 = ttk.LabelFrame(page, text="① 扫描目标(可混选分区 / 文件夹 / 单个文件)", padding=6)
        box1.grid(row=0, column=0, sticky="ew")
        box1.columnconfigure(0, weight=1)

        bar = ttk.Frame(box1)
        bar.grid(row=0, column=0, sticky="ew")
        ttk.Label(bar, text="分区:").pack(side="left")
        self.cmb_drive = ttk.Combobox(bar, textvariable=self.var_drive, width=42,
                                      state="readonly")
        self.cmb_drive.pack(side="left", padx=4)
        ttk.Button(bar, text="加入目标", command=self._add_drive_target).pack(side="left", padx=2)
        ttk.Button(bar, text="刷新分区", command=self._refresh_drives).pack(side="left", padx=2)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(bar, text="添加文件夹", command=self._add_folder_target).pack(side="left", padx=2)
        ttk.Button(bar, text="添加文件", command=self._add_file_target).pack(side="left", padx=2)
        ttk.Button(bar, text="移除选中", command=self._remove_target).pack(side="left", padx=2)
        ttk.Button(bar, text="清空", command=self._clear_targets).pack(side="left", padx=2)

        cols = ("kind", "path")
        self.tree_targets = ttk.Treeview(box1, columns=cols, show="headings", height=4)
        self.tree_targets.heading("kind", text="类型")
        self.tree_targets.heading("path", text="路径")
        self.tree_targets.column("kind", width=80, anchor="center", stretch=False)
        self.tree_targets.column("path", width=900, anchor="w")
        self.tree_targets.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        scroll = ttk.Scrollbar(box1, orient="horizontal", command=self.tree_targets.xview)
        scroll.grid(row=2, column=0, sticky="ew", pady=(2, 0))
        self.tree_targets.configure(xscrollcommand=scroll.set)

        # ② 扫描选项 --------------------------------------------------- #
        box2 = ttk.LabelFrame(page, text="② 扫描", padding=6)
        box2.grid(row=1, column=0, sticky="ew", pady=(8, 0))

        left = ttk.Frame(box2)
        left.pack(side="left", fill="x", expand=True)

        opts = ttk.Frame(left)
        opts.pack(side="top", anchor="w")
        ttk.Checkbutton(opts, text="递归子目录", variable=self.var_recursive).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(opts, text="跳过系统目录", variable=self.var_skip_system).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(opts, text="移除阻止标记(Protected View)",
                        variable=self.var_unblock).pack(side="left", padx=(0, 14))
        self.btn_scan = ttk.Button(opts, text="开始扫描", command=self._start_scan)
        self.btn_scan.pack(side="left", padx=(0, 6))
        self.btn_scan_cancel = ttk.Button(opts, text="取消", command=self._cancel, state="disabled")
        self.btn_scan_cancel.pack(side="left")

        kopts = ttk.Frame(left)
        kopts.pack(side="top", anchor="w", pady=(6, 0))
        ttk.Label(kopts, text="搜索内核:").pack(side="left")
        self.cmb_kernel = ttk.Combobox(
            kopts, textvariable=self.var_kernel, width=34, state="readonly",
            values=[KERNEL_LABELS[KERNEL_AUTO], KERNEL_LABELS[KERNEL_EVERYTHING],
                    KERNEL_LABELS[KERNEL_WINDOWS_INDEX], KERNEL_LABELS[KERNEL_WALK]])
        self.cmb_kernel.pack(side="left", padx=4)
        ttk.Label(kopts, text="← 索引内核毫秒级;直接遍历最慢").pack(side="left")

        ttk.Label(box2, textvariable=self.var_scan_info).pack(side="right")

        # ③ 文件清单 --------------------------------------------------- #
        box3 = ttk.LabelFrame(page, text="③ 待转换文件清单", padding=6)
        box3.grid(row=3, column=0, sticky="nsew", pady=(8, 0))
        box3.columnconfigure(0, weight=1)
        box3.rowconfigure(0, weight=1)

        cols3 = ("name", "dir", "size", "mtime", "status", "message")
        self.tree_files = ttk.Treeview(box3, columns=cols3, show="headings", selectmode="extended")
        heads = [("name", "文件名", 220, "w"), ("dir", "所在目录", 300, "w"),
                 ("size", "大小", 90, "e"), ("mtime", "修改时间", 140, "center"),
                 ("status", "状态", 80, "center"), ("message", "说明", 260, "w")]
        for key, title, width, anchor in heads:
            self.tree_files.heading(key, text=title)
            self.tree_files.column(key, width=width, anchor=anchor)
        self.tree_files.tag_configure(STATUS_SUCCESS, foreground=STATUS_COLORS[STATUS_SUCCESS])
        self.tree_files.tag_configure(STATUS_SKIPPED, foreground=STATUS_COLORS[STATUS_SKIPPED])
        self.tree_files.tag_configure(STATUS_FAILED, foreground=STATUS_COLORS[STATUS_FAILED])
        self.tree_files.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box3, orient="vertical", command=self.tree_files.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree_files.configure(yscrollcommand=sb.set)

        # ④ 转换选项 --------------------------------------------------- #
        box4 = ttk.LabelFrame(page, text="④ 转换选项", padding=6)
        box4.grid(row=4, column=0, sticky="ew", pady=(8, 0))

        left_col = ttk.Frame(box4)
        left_col.pack(side="left", fill="x", expand=True)

        line1 = ttk.Frame(left_col)
        line1.pack(side="top", anchor="w")
        ttk.Label(line1, text="输出位置:").pack(side="left")
        ttk.Radiobutton(line1, text="与源文件同目录", value=True,
                        variable=self.var_out_same).pack(side="left", padx=(4, 10))
        ttk.Radiobutton(line1, text="指定目录:", value=False,
                        variable=self.var_out_same).pack(side="left")
        self.ent_out = ttk.Entry(line1, textvariable=self.var_out_dir, width=38)
        self.ent_out.pack(side="left", padx=4)
        ttk.Button(line1, text="浏览", command=self._pick_output_dir).pack(side="left")

        line3 = ttk.Frame(left_col)
        line3.pack(side="top", anchor="w", pady=(6, 0))
        ttk.Label(line3, text="转换成功后处理原 .docm:").pack(side="left")
        self.cmb_cleanup = ttk.Combobox(
            line3, textvariable=self.var_cleanup, width=30, state="readonly",
            values=[CLEANUP_LABELS[CLEANUP_NONE], CLEANUP_LABELS[CLEANUP_RECYCLE],
                    CLEANUP_LABELS[CLEANUP_DELETE]])
        self.cmb_cleanup.pack(side="left", padx=4)
        ttk.Label(line3, text="← 仅对转换成功的文件生效,跳过/失败的一律保留").pack(side="left")

        line2 = ttk.Frame(box4)
        line2.pack(side="right")
        ttk.Label(line2, text="目标已存在:").pack(side="left")
        self.cmb_overwrite = ttk.Combobox(line2, textvariable=self.var_overwrite,
                                          width=24, state="readonly",
                                          values=[OVERWRITE_LABELS[OVERWRITE_SKIP],
                                                  OVERWRITE_LABELS[OVERWRITE_REPLACE],
                                                  OVERWRITE_LABELS[OVERWRITE_RENAME]])
        self.cmb_overwrite.pack(side="left", padx=4)
        self.btn_convert = ttk.Button(line2, text="开始转换", command=self._start_convert)
        self.btn_convert.pack(side="left", padx=(12, 6))
        self.btn_cancel = ttk.Button(line2, text="取消", command=self._cancel, state="disabled")
        self.btn_cancel.pack(side="left", padx=(0, 6))
        ttk.Button(line2, text="导出日志", command=self._export_log).pack(side="left")

        # ⑤ 进度 ------------------------------------------------------- #
        box5 = ttk.LabelFrame(page, text="⑤ 进度", padding=6)
        box5.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        box5.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(box5, mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew")
        top = ttk.Frame(box5)
        top.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        ttk.Label(top, textvariable=self.var_status).pack(side="left")
        ttk.Label(top, textvariable=self.var_counters).pack(side="right")

        bottom = ttk.Frame(box5)
        bottom.grid(row=2, column=0, sticky="ew")
        ttk.Label(bottom, textvariable=self.var_speed).pack(side="right")

        # ⑥ 日志 ------------------------------------------------------- #
        box6 = ttk.LabelFrame(page, text="⑥ 日志", padding=6)
        box6.grid(row=6, column=0, sticky="ew", pady=(8, 0))
        box6.columnconfigure(0, weight=1)
        self.log_text = ScrolledText(box6, height=8, wrap="none", font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky="ew")
        self.log_text.configure(state="disabled")

        return page

    # ---------------------------------------------------------------- 测试页 #
    def _build_test_tab(self, parent) -> ttk.Frame:
        page = ttk.Frame(parent, padding=8)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(2, weight=1)

        box1 = ttk.LabelFrame(page, text="测试工作目录", padding=6)
        box1.grid(row=0, column=0, sticky="ew")
        line = ttk.Frame(box1)
        line.pack(side="left")
        ttk.Label(line, text="目录:").pack(side="left")
        self.ent_test_dir = ttk.Entry(line, textvariable=self.test_dir_var(), width=1)
        self.ent_test_dir.pack(side="left", padx=4, fill="x", expand=True)
        ttk.Button(line, text="浏览", command=self._pick_test_dir).pack(side="left")

        box2 = ttk.Frame(page)
        box2.grid(row=1, column=0, sticky="ew", pady=8)
        self.btn_gen = ttk.Button(box2, text="① 生成测试 DOCM", command=self._start_gen_test)
        self.btn_gen.pack(side="left", padx=(0, 6))
        self.btn_conv_test = ttk.Button(box2, text="② 转换为 DOCX",
                                         command=self._start_convert_test, state="disabled")
        self.btn_conv_test.pack(side="left", padx=(0, 6))
        self.btn_verify = ttk.Button(box2, text="③ 机器校验",
                                     command=self._start_verify, state="disabled")
        self.btn_verify.pack(side="left", padx=(0, 6))
        self.btn_manual = ttk.Button(box2, text="④ 人工复核(在 Word 中并排打开)",
                                      command=self._start_manual, state="disabled")
        self.btn_manual.pack(side="left", padx=(0, 6))
        ttk.Label(box2, textvariable=self.var_test_status).pack(side="right")

        box3 = ttk.LabelFrame(page, text="校验结果", padding=6)
        box3.grid(row=2, column=0, sticky="nsew")
        box3.columnconfigure(0, weight=1)
        box3.rowconfigure(0, weight=1)

        cols = ("name", "result", "detail")
        self.tree_checks = ttk.Treeview(box3, columns=cols, show="headings")
        self.tree_checks.heading("name", text="检查项")
        self.tree_checks.heading("result", text="结果")
        self.tree_checks.heading("detail", text="说明")
        self.tree_checks.column("name", width=260, anchor="w", stretch=False)
        self.tree_checks.column("result", width=90, anchor="center", stretch=False)
        self.tree_checks.column("detail", width=760, anchor="w")
        self.tree_checks.tag_configure("fail", foreground="#c5221f")
        self.tree_checks.tag_configure("warn", foreground="#8a6d00")
        self.tree_checks.tag_configure("ok", foreground="#137333")
        self.tree_checks.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box3, orient="vertical", command=self.tree_checks.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree_checks.configure(yscrollcommand=sb.set)

        box4 = ttk.LabelFrame(page, text="详情", padding=6)
        box4.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        box4.columnconfigure(0, weight=1)
        self.test_detail = ScrolledText(box4, height=10, wrap="word", font=("Consolas", 9))
        self.test_detail.grid(row=0, column=0, sticky="ew")
        self.test_detail.configure(state="disabled")

        return page

    def test_dir_var(self) -> tk.StringVar:
        if not hasattr(self, "_test_dir_var"):
            self._test_dir_var = tk.StringVar(value=self.test_dir)
        return self._test_dir_var

    # ===================================================================== #
    # 目标管理
    # ===================================================================== #
    def _refresh_drives(self) -> None:
        drives = scanner.list_drives()
        values = ["— 全部分区(整盘扫描) —"]
        for drive in drives:
            values.append(drive.display)
        self.cmb_drive.configure(values=values)
        if values:
            self.cmb_drive.current(0)
        self._drives = drives

    def _add_drive_target(self) -> None:
        index = self.cmb_drive.current()
        if index is None or index < 0:
            return
        if index == 0:
            for drive in getattr(self, "_drives", []):
                self._append_target("drive", drive.path)
            self._log("已加入全部分区")
            return
        drive = self._drives[index - 1]
        self._append_target("drive", drive.path)
        self._log(f"已加入分区 {drive.path}")

    def _append_target(self, kind: str, path: str) -> None:
        path = os.path.abspath(path)
        for item in self.targets:
            if item["kind"] == kind and item["path"].lower() == path.lower():
                self._log(f"目标已存在,跳过: {path}")
                return
        self.targets.append({"kind": kind, "path": path})
        self.tree_targets.insert("", "end", values=(KIND_LABELS[kind], path))

    def _add_folder_target(self) -> None:
        path = filedialog.askdirectory(title="选择要扫描的文件夹")
        if path:
            self._append_target("folder", path)

    def _add_file_target(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择 .docm 文件(可多选)",
            filetypes=[("Word 宏文档", "*.docm"), ("所有文件", "*.*")])
        for path in paths:
            self._append_target("file", path)

    def _remove_target(self) -> None:
        for item in self.tree_targets.selection():
            index = self.tree_targets.index(item)
            self.targets.pop(index)
            self.tree_targets.delete(item)

    def _clear_targets(self) -> None:
        self.targets.clear()
        self.tree_targets.delete(*self.tree_targets.get_children())

    def _pick_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.var_out_dir.set(path)
            self.var_out_same.set(False)

    def _pick_test_dir(self) -> None:
        path = filedialog.askdirectory(title="选择测试工作目录")
        if path:
            self.test_dir = path
            self.test_dir_var().set(path)

    # ===================================================================== #
    # 扫描
    # ===================================================================== #
    def _start_scan(self) -> None:
        if not self.targets:
            messagebox.showinfo("提示", "请先添加至少一个扫描目标。")
            return
        self.cancel_event.clear()
        self.files.clear()
        self.file_records.clear()
        self.row_ids.clear()
        self.tree_files.delete(*self.tree_files.get_children())
        self.progress.configure(value=0, maximum=100)
        self.var_scan_info.set("正在扫描 …")
        self._set_busy(True)

        roots = [t["path"] for t in self.targets if t["kind"] in ("drive", "folder")]
        singles = [t["path"] for t in self.targets if t["kind"] == "file"]
        recursive = self.var_recursive.get()
        kernel = KERNEL_BY_LABEL.get(self.var_kernel.get(), KERNEL_AUTO)

        def worker():
            def progress(current_dir: str, count: int) -> None:
                self._post("scan_progress", (current_dir, count))
            started = time.time()
            outcome = search_kernel.find_docm_files(
                roots, recursive=recursive, kernel=kernel,
                cancel=self.cancel_event.is_set, progress=progress)
            files = list(outcome.files)
            for path in singles:
                if (path.lower().endswith(".docm") and os.path.isfile(path)
                        and not os.path.basename(path).startswith("~$")):
                    files.append(path)
            unique, seen = [], set()
            for path in files:
                key = os.path.normcase(os.path.abspath(path))
                if key in seen:
                    continue
                seen.add(key)
                unique.append(path)
            unique.sort(key=lambda p: p.lower())
            self._post("scan_done", (unique, outcome.errors, outcome.kernel_name,
                                     time.time() - started, outcome.stale_skipped))

        self._run_bg(worker, "扫描")

    def _on_scan_done(self, files, errors, kernel_name, elapsed, stale=0) -> None:
        roots = {os.path.normcase(t["path"]): t["path"]
                 for t in self.targets if t["kind"] != "file"}
        display_limit = 20000          # 表格最多渲染这么多行,避免界面卡死
        shown = 0
        for path in files:
            info = scanner.file_info(path)
            root = ""
            for base in roots.values():
                if path.lower().startswith(base.lower() + os.sep):
                    if len(base) > len(root):
                        root = base
            record = {"path": path, "size": info["size"], "mtime": info["mtime"],
                      "status": STATUS_PENDING, "message": "", "root": root}
            self.files.append(record)
            self.file_records[path] = record
            if shown < display_limit:
                iid = self.tree_files.insert("", "end", values=(
                    os.path.basename(path), os.path.dirname(path),
                    scanner.human_size(info["size"]), info["mtime"],
                    STATUS_PENDING, ""))
                self.row_ids[path] = iid
                shown += 1
        self.tree_files.update_idletasks()
        self.var_scan_info.set(
            f"找到 {len(files):,} 个 .docm · 内核 {kernel_name} · 用时 {elapsed:.2f}s")
        self._log(f"扫描完成:共 {len(files):,} 个 .docm"
                  f"(内核 {kernel_name},用时 {elapsed:.2f}s)")
        if stale:
            self._log(f"  索引中已失效的条目已自动剔除:{stale} 个")
        if len(files) > display_limit:
            self._log(f"  清单仅显示前 {display_limit:,} 行,"
                      f"转换仍会覆盖全部 {len(files):,} 个文件")
        for line in errors[:10]:
            self._log(f"  内核提示: {line}")
        if len(errors) > 10:
            self._log(f"  …… 另有 {len(errors) - 10} 条内核提示")
        if not files:
            messagebox.showinfo("扫描结果", "未找到任何 .docm 文件。")

    # ===================================================================== #
    # 转换
    # ===================================================================== #
    def _start_convert(self) -> None:
        pending = [f for f in self.files if f["status"] in (STATUS_PENDING, STATUS_FAILED)]
        if not pending:
            messagebox.showinfo("提示", "没有待转换的文件,请先扫描。")
            return
        output_root = "" if self.var_out_same.get() else self.var_out_dir.get().strip()
        if not self.var_out_same.get() and not output_root:
            messagebox.showinfo("提示", "请选择输出目录,或改为“与源文件同目录”。")
            return
        if output_root and not os.path.isdir(output_root):
            try:
                os.makedirs(output_root, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("错误", f"无法创建输出目录: {exc}")
                return

        overwrite = OVERWRITE_BY_LABEL.get(self.var_overwrite.get(), OVERWRITE_SKIP)
        cleanup = CLEANUP_BY_LABEL.get(self.var_cleanup.get(), CLEANUP_NONE)

        if cleanup == CLEANUP_DELETE:
            if not messagebox.askyesno(
                    "确认永久删除源文件",
                    "你选择了「永久删除原 .docm」。\n\n"
                    "转换成功后,对应的原 .docm 将被**永久删除**,\n"
                    "不会进入回收站,无法恢复。\n\n"
                    f"本次待处理 {len(pending)} 个文件。\n"
                    "如果只是想去掉源文件,建议改选「移到回收站」。\n\n"
                    "确定要继续吗?"):
                return
        elif cleanup == CLEANUP_RECYCLE:
            if not messagebox.askyesno(
                    "确认处理源文件",
                    "你选择了「移到回收站」。\n\n"
                    "转换成功后,对应的原 .docm 会被移入 Windows 回收站,\n"
                    "需要时可以还原。\n\n"
                    f"本次待处理 {len(pending)} 个文件。\n\n继续吗?"):
                return

        self.cancel_event.clear()
        self.progress.configure(value=0, maximum=len(pending))
        self.counters = {STATUS_SUCCESS: 0, STATUS_SKIPPED: 0, STATUS_FAILED: 0}
        self.var_counters.set("成功 0　跳过 0　失败 0　总计 %d" % len(pending))
        self.var_speed.set("用时 00:00　速度 -")
        self._set_busy(True)
        started = time.time()

        # Tk 变量只能在主线程读取,先取出再交给后台线程
        unblock = self.var_unblock.get()

        items = [ConvertItem(src=f["path"], root=f.get("root", ""),
                             dst=plan_output(f["path"], f.get("root", ""), output_root))
                 for f in pending]

        def worker():
            def on_item(item: ConvertItem, index: int, total: int) -> None:
                self._post("item", (item, index, total))
            summary = convert_batch(
                items, overwrite=overwrite, output_root=output_root,
                cancel=self.cancel_event.is_set, on_item=on_item,
                log=lambda m: self._post("log", m), unblock=unblock,
                cleanup=cleanup)
            self._post("convert_done", (summary, started))

        self._run_bg(worker, "转换")

    def _on_item(self, item: ConvertItem, index: int, total: int) -> None:
        record = self.file_records.get(item.src)
        if record is not None:
            record["status"] = item.status
            record["message"] = item.message
        iid = self.row_ids.get(item.src)
        if iid:
            try:
                self.tree_files.item(
                    iid, values=(
                        os.path.basename(item.src), os.path.dirname(item.src),
                        self._size_text(item.src), self._mtime_text(item.src),
                        item.status, item.message))
                self.tree_files.item(iid, tags=(item.status,))
            except tk.TclError:
                pass
        self.progress.configure(value=index, maximum=total)
        self.var_status.set(f"[{index}/{total}] {os.path.basename(item.src)} — {item.message}")
        elapsed = max(0.001, time.time() - self._convert_started)
        speed = index / elapsed * 60
        self.counters[item.status] = self.counters.get(item.status, 0) + 1
        self.var_counters.set(f"成功 {self.counters[STATUS_SUCCESS]}　"
                              f"跳过 {self.counters[STATUS_SKIPPED]}　"
                              f"失败 {self.counters[STATUS_FAILED]}　总计 {total}")
        self.var_speed.set(f"用时 {self._fmt_time(elapsed)}　速度 {speed:.1f} 个/分")
        self._log(f"[{index}/{total}] {item.status} — {os.path.basename(item.src)}"
                  + (f" — {item.message}" if item.message else ""))

    @staticmethod
    def _size_text(path: str) -> str:
        try:
            return scanner.human_size(os.path.getsize(path))
        except OSError:
            return "-"

    @staticmethod
    def _mtime_text(path: str) -> str:
        try:
            return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(os.path.getmtime(path)))
        except OSError:
            return "-"

    @staticmethod
    def _fmt_time(seconds: float) -> str:
        seconds = int(seconds)
        return f"{seconds // 60:02d}:{seconds % 60:02d}"

    def _on_convert_done(self, summary, started: float) -> None:
        self.var_status.set(
            f"转换结束:成功 {summary.success}　跳过 {summary.skipped}　失败 {summary.failed}"
            + ("(已取消)" if summary.cancelled else ""))
        self.var_speed.set(f"用时 {self._fmt_time(summary.elapsed)}　"
                           f"平均 {summary.done / max(0.001, summary.elapsed) * 60:.1f} 个/分")
        self._log("=" * 60)
        self._log(f"转换完成:总计 {summary.total}　成功 {summary.success}　"
                  f"跳过 {summary.skipped}　失败 {summary.failed}　"
                  f"耗时 {summary.elapsed:.1f}s")
        if summary.removed:
            self._log(f"源文件处理:已处理 {summary.removed} 个(仅转换成功的文件)")
        if summary.cleanup_failed:
            self._log(f"源文件处理失败:{summary.cleanup_failed} 个,源文件已保留")
        for path, message in summary.errors[:20]:
            self._log(f"  失败: {path} — {message}")
        if summary.cancelled:
            messagebox.showinfo("转换结束", "转换已被用户取消。")
        elif summary.failed:
            messagebox.showwarning("转换结束",
                                   f"完成,但有 {summary.failed} 个文件失败,详见日志。")
        else:
            extra = ""
            if summary.removed:
                extra = f"\n\n已按设置处理 {summary.removed} 个原 .docm。"
            messagebox.showinfo("转换结束",
                                f"全部完成!成功 {summary.success} 个,跳过 {summary.skipped} 个。"
                                f"{extra}")

    # ===================================================================== #
    # 测试与校验
    # ===================================================================== #
    def _start_gen_test(self) -> None:
        directory = self.test_dir_var().get().strip() or self.test_dir
        try:
            os.makedirs(directory, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("错误", f"无法创建测试目录: {exc}")
            return
        self.test_dir = directory
        docm_path = os.path.join(directory, "转换测试用例.docm")
        self._set_busy(True)
        self.var_test_status.set("正在生成测试文档 …")

        def worker():
            session = WordSession(visible=False)
            try:
                session.start()
                info = create_test_docm(session.app, docm_path)
                self._post("test_gen_done", info)
            finally:
                session.close()

        self._run_bg(worker, "生成测试文档")

    def _on_test_gen_done(self, info) -> None:
        self.test_docm = info["docm"]
        self.test_docx = os.path.splitext(self.test_docm)[0] + ".docx"
        self.var_test_status.set(f"已生成:{os.path.basename(self.test_docm)}"
                                 f"(含 {len(info['created'])} 类元素)")
        self._log(f"[测试] 已生成 {self.test_docm}")
        for name in info["created"]:
            self._log(f"  · {name}")
        for note in info["notes"]:
            self._log(f"  [警告] {note}")
        self.btn_conv_test.configure(state="normal")
        self._append_test_detail(
            f"测试文档已生成\n路径:{self.test_docm}\n"
            f"写入元素:{len(info['created'])} 类\n"
            + "\n".join(f"  · {n}" for n in info["created"]))

    def _start_convert_test(self) -> None:
        if not self.test_docm:
            return
        self._set_busy(True)
        self.var_test_status.set("正在转换测试文档 …")

        def worker():
            item = ConvertItem(src=self.test_docm, dst=self.test_docx)
            summary = convert_batch([item], overwrite=OVERWRITE_REPLACE)
            self._post("test_conv_done", (item, summary))

        self._run_bg(worker, "测试转换")

    def _on_test_conv_done(self, item, summary) -> None:
        if item.status == STATUS_SUCCESS:
            self.var_test_status.set(f"转换成功:{os.path.basename(self.test_docx)}")
            self._log(f"[测试] 转换成功 → {self.test_docx}")
            self.btn_verify.configure(state="normal")
            self.btn_manual.configure(state="normal")
        else:
            self.var_test_status.set(f"转换失败:{item.message}")
            self._log(f"[测试] 转换失败:{item.message}")
        self._append_test_detail(f"转换结果:{item.status}\n{item.message}\n"
                                 f"产物:{self.test_docx}")

    def _start_verify(self) -> None:
        if not self.test_docx or not os.path.isfile(self.test_docx):
            messagebox.showinfo("提示", "请先生成测试文档并完成转换。")
            return
        self._set_busy(True)
        self.var_test_status.set("正在机器校验 …")
        self.tree_checks.delete(*self.tree_checks.get_children())

        def worker():
            report = verify_conversion(self.test_docm, self.test_docx)
            self._post("verify_done", report)

        self._run_bg(worker, "机器校验")

    def _on_verify_done(self, report) -> None:
        for check in report.checks:
            tag = {"fail": "fail", "warn": "warn"}.get(check.level, "ok")
            self.tree_checks.insert("", "end", values=(check.name, check.status, check.detail),
                                    tags=(tag,))
        self.var_test_status.set("校验完成:" + report.summary)
        lines = []
        for check in report.checks:
            lines.append(f"{check.status}  {check.name}")
            lines.append(f"    {check.detail}")
        lines.append("")
        lines.append(f"结论:{'全部通过 —— 转换无损' if report.ok else '存在失败项'}({report.summary})")
        self._append_test_detail("\n".join(lines))
        for check in report.checks:
            self._log(f"[校验] {check.status} {check.name} — {check.detail}")
        if report.ok:
            messagebox.showinfo("校验结果",
                                f"全部通过!转换无损。\n\n{report.summary}\n\n"
                                "可再点“人工复核”在 Word 中并排打开两个文件肉眼确认。")
        else:
            messagebox.showerror("校验结果", f"存在失败项,请查看明细。\n\n{report.summary}")

    def _start_manual(self) -> None:
        if not self.test_docx or not os.path.isfile(self.test_docx):
            messagebox.showinfo("提示", "请先完成转换。")
            return
        self._set_busy(True)
        self.var_test_status.set("正在打开 Word 供人工比对 …")

        def worker():
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = True
            app.DisplayAlerts = 0
            try:
                app.AutomationSecurity = 3
            except Exception:
                pass
            for path in (self.test_docm, self.test_docx):
                try:
                    app.Documents.Open(FileName=path, ReadOnly=True,
                                       AddToRecentFiles=False, Visible=True)
                except Exception as exc:
                    self._post("log", f"[人工复核] 打开失败 {path}: {exc}")
            self._post("manual_done", None)
            # 刻意不退出 Word:留给用户肉眼比对,关闭 Word 窗口即结束
            pythoncom.CoUninitialize()

        self._run_bg(worker, "人工复核")

    def _on_manual_done(self, _) -> None:
        self.var_test_status.set("已在 Word 中打开 .docm 与 .docx,请肉眼比对后关闭 Word")
        self._append_test_detail(
            "已在 Word 中并排打开:\n"
            f"  源文件:{self.test_docm}\n"
            f"  产物  :{self.test_docx}\n\n"
            "人工核对要点:\n"
            "  1. 标题层级与正文字体、字号是否一致\n"
            "  2. 项目符号 / 编号列表是否正常\n"
            "  3. 两张表格的行列与内容是否完整\n"
            "  4. 图片与矩形形状是否显示正常\n"
            "  5. 超链接可点击、日期/页码域是否正确\n"
            "  6. 页眉页脚、脚注、第二页分页是否正常\n"
            "  7. 打开时是否出现“Word 发现无法读取的内容”提示"
            "(若出现即为损坏,请回到日志查看失败原因)")

    # ===================================================================== #
    # 线程与消息
    # ===================================================================== #
    def _post(self, kind: str, payload) -> None:
        self.queue.put((kind, payload))

    def _run_bg(self, func, name: str) -> None:
        def runner():
            try:
                func()
            except Exception as exc:
                self._post("error", f"{name}失败:{exc}")
            finally:
                self._post("done", name)

        self._convert_started = time.time()
        threading.Thread(target=runner, daemon=True).start()

    def _poll(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                self._handle(kind, payload)
        except queue.Empty:
            pass
        except tk.TclError:
            return                      # 窗口已销毁,停止轮询
        except Exception as exc:       # 单条消息出错不应中断整个界面
            self._log(f"[界面] 处理消息时出错: {exc}")
        try:
            self._poll_id = self.after(120, self._poll)
        except tk.TclError:
            pass

    def _handle(self, kind: str, payload) -> None:
        if kind == "log":
            self._log(str(payload))
        elif kind == "scan_progress":
            current, count = payload
            tail = os.path.basename(current.rstrip("\\/")) or current
            self.var_scan_info.set(f"正在扫描 … 已找到 {count:,} 个 .docm(当前:{tail})")
        elif kind == "scan_done":
            self._on_scan_done(*payload)
        elif kind == "item":
            self._on_item(*payload)
        elif kind == "convert_done":
            summary, started = payload
            self._on_convert_done(summary, started)
        elif kind == "test_gen_done":
            self._on_test_gen_done(payload)
        elif kind == "test_conv_done":
            self._on_test_conv_done(*payload)
        elif kind == "verify_done":
            self._on_verify_done(payload)
        elif kind == "manual_done":
            self._on_manual_done(payload)
        elif kind == "error":
            self._log(str(payload))
            messagebox.showerror("出错了", str(payload))
        elif kind == "done":
            self._set_busy(False)

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.btn_scan.configure(state=state)
        self.btn_convert.configure(state=state)
        self.btn_gen.configure(state=state)
        self.btn_conv_test.configure(state=state if self.test_docm else "disabled")
        self.btn_verify.configure(state=state if self.test_docx else "disabled")
        self.btn_manual.configure(state=state if self.test_docx else "disabled")
        self.btn_scan_cancel.configure(state="normal" if busy else "disabled")
        self.btn_cancel.configure(state="normal" if busy else "disabled")

    def _cancel(self) -> None:
        if self.busy:
            self.cancel_event.set()
            self._log("已请求取消,正在停止 …")
            self.var_status.set("正在取消 …")

    # ===================================================================== #
    # 日志
    # ===================================================================== #
    def _log(self, text: str) -> None:
        stamp = time.strftime("%H:%M:%S")
        line = f"[{stamp}] {text}"
        self.log_lines.append(line)
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _append_test_detail(self, text: str) -> None:
        self.test_detail.configure(state="normal")
        self.test_detail.insert("end", text + "\n\n")
        self.test_detail.see("end")
        self.test_detail.configure(state="disabled")

    def _export_log(self) -> None:
        if not self.log_lines:
            messagebox.showinfo("日志", "暂无日志内容。")
            return
        path = filedialog.asksaveasfilename(
            title="导出日志", defaultextension=".log",
            filetypes=[("日志文件", "*.log"), ("文本文件", "*.txt")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("\n".join(self.log_lines))
        except OSError as exc:
            messagebox.showerror("错误", f"导出失败:{exc}")
            return
        self._log(f"日志已导出:{path}")
        messagebox.showinfo("完成", f"日志已导出到:\n{path}")

    # ===================================================================== #
    def _on_close(self) -> None:
        if self.busy:
            if not messagebox.askyesno("退出", "仍有任务在运行,确定要退出吗?"):
                return
            self.cancel_event.set()
        if self._poll_id is not None:
            try:
                self.after_cancel(self._poll_id)
            except tk.TclError:
                pass
            self._poll_id = None
        try:
            search_kernel.shutdown_everything()      # 关掉我们启动的内置 Everything
        except Exception:
            pass
        self.destroy()


def main() -> int:
    app = ConverterApp()
    app.protocol("WM_DELETE_WINDOW", app._on_close)
    app.mainloop()
    return 0
