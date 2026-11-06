# -*- coding: utf-8 -*-
"""命令行入口:扫描 / 转换 / 自检。

GUI 之外的自动化用法,也用于在没有交互环境时验证转换链路。
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from typing import Sequence

from .constants import DOCX_SUFFIX, OVERWRITE_LABELS
from .engine import (
    ConvertItem, STATUS_FAILED, STATUS_SKIPPED, STATUS_SUCCESS,
    WordSession, convert_batch, plan_output,
)
from .search_kernel import (
    KERNEL_AUTO, KERNEL_EVERYTHING, KERNEL_LABELS, KERNEL_WALK,
    KERNEL_WINDOWS_INDEX, find_docm_files,
)
from .testgen import create_test_docm
from .verify import verify_conversion


def _print(msg: str = "") -> None:
    try:
        print(msg)
    except UnicodeEncodeError:                      # 老终端 GBK 兜底
        print(msg.encode("utf-8", "replace").decode("ascii", "replace"))


# --------------------------------------------------------------------------- #
# scan
# --------------------------------------------------------------------------- #
def _split_targets(paths):
    """把命令行路径拆成 (目录根, 单文件)。"""
    roots, singles = [], []
    for path in paths:
        path = os.path.abspath(path)
        if os.path.isfile(path):
            singles.append(path)
        elif os.path.isdir(path):
            roots.append(path)
        else:
            _print(f"跳过无效路径: {path}")
    return roots, singles


def _collect(paths, kernel, recursive):
    """按指定内核收集 .docm。"""
    roots, singles = _split_targets(paths)
    outcome = find_docm_files(roots, recursive=recursive, kernel=kernel)
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
    return unique, outcome


# --------------------------------------------------------------------------- #
# scan
# --------------------------------------------------------------------------- #
def cmd_scan(args) -> int:
    files, outcome = _collect(args.paths, args.kernel, not args.no_recursive)
    _print(f"搜索内核: {outcome.kernel_name}  用时 {outcome.elapsed:.2f}s")
    _print(f"找到 .docm: {len(files)} 个")
    for item in files:
        _print(f"  {item}")
    if outcome.errors:
        _print(f"\n内核提示 {len(outcome.errors)} 条(前 10 条):")
        for line in outcome.errors[:10]:
            _print(f"  {line}")
    return 0


# --------------------------------------------------------------------------- #
# convert
# --------------------------------------------------------------------------- #
def cmd_convert(args) -> int:
    if args.out:
        args.out = os.path.abspath(args.out)
        try:
            os.makedirs(args.out, exist_ok=True)
        except OSError as exc:
            _print(f"错误:无法创建输出目录 {args.out}: {exc}")
            return 1

    if args.delete_source != "none" and not args.yes:
        tip = ("转换成功后,原 .docm 会被移到回收站(可恢复)。"
               if args.delete_source == "recycle"
               else "转换成功后,原 .docm 会被永久删除,无法恢复!")
        _print("")
        _print(f"!! {tip}")
        try:
            answer = input("   确认继续?[y/N] ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            answer = ""
        if answer not in ("y", "yes"):
            _print("   已取消。")
            return 1

    roots_list, _ = _split_targets(args.paths)
    files, outcome = _collect(args.paths, args.kernel, not args.no_recursive)
    if not files:
        _print("未找到任何 .docm 文件。")
        return 1
    _print(f"搜索内核: {outcome.kernel_name}  用时 {outcome.elapsed:.2f}s")

    roots: dict = {}
    for path in roots_list:
        roots[os.path.normcase(path)] = path

    items = []
    for path in files:
        root = ""
        for base in roots.values():
            if path.lower().startswith(base.lower() + os.sep):
                root = base
                break
        items.append(ConvertItem(src=path, root=root))

    started = time.time()

    def on_item(item, index, total):
        mark = {"成功": "+", "跳过": "=", "失败": "!"}.get(item.status, "?")
        _print(f"[{index}/{total}] {mark} {item.name} — {item.message}")

    summary = convert_batch(
        items,
        overwrite=args.overwrite,
        output_root=args.out or "",
        log=lambda m: _print(f"  · {m}"),
        on_item=on_item,
        cleanup=args.delete_source,
    )
    _print("")
    _print("=" * 56)
    _print(f"总计 {summary.total} | 成功 {summary.success} | "
           f"跳过 {summary.skipped} | 失败 {summary.failed} | "
           f"耗时 {summary.elapsed:.1f}s")
    if summary.removed:
        _print(f"源文件已按设置处理:{summary.removed} 个")
    if summary.cleanup_failed:
        _print(f"源文件处理失败:{summary.cleanup_failed} 个(源文件已保留)")
    if summary.errors:
        _print("\n失败明细:")
        for path, message in summary.errors:
            _print(f"  {path}\n    {message}")
    return 0 if summary.failed == 0 else 2


# --------------------------------------------------------------------------- #
# selftest
# --------------------------------------------------------------------------- #
def cmd_selftest(args) -> int:
    out_dir = os.path.abspath(args.dir)
    os.makedirs(out_dir, exist_ok=True)
    docm_path = os.path.join(out_dir, "转换测试用例.docm")
    docx_path = os.path.splitext(docm_path)[0] + DOCX_SUFFIX
    if os.path.exists(docx_path):
        os.remove(docx_path)

    _print("=" * 56)
    _print("DOCM → DOCX 无损转换自检")
    _print("=" * 56)
    _print(f"输出目录: {out_dir}\n")

    session = WordSession(visible=False)
    _print("[1/4] 启动 Word 并生成测试 .docm …")
    session.start()
    try:
        info = create_test_docm(session.app, docm_path)
        created = info["created"]
        _print(f"      已写入 {len(created)} 类元素:")
        for name in created:
            _print(f"        · {name}")
        for note in info["notes"]:
            _print(f"      [警告] {note}")

        _print("\n[2/4] 执行转换(docm → docx) …")
        item = ConvertItem(src=docm_path, dst=docx_path)
        summary = convert_batch([item], overwrite="overwrite", log=lambda m: _print(f"      {m}"))
        if item.status != STATUS_SUCCESS:
            _print(f"      转换失败: {item.message}")
            return 2
        _print(f"      转换成功 → {docx_path}")
    finally:
        session.close()

    _print("\n[3/4] 机器校验(docm ↔ docx 逐项比对) …")
    report = verify_conversion(docm_path, docx_path)
    for check in report.checks:
        _print(f"      {check.status}  {check.name}")
        if check.detail:
            _print(f"            {check.detail}")

    _print("\n[4/4] 结论")
    if report.ok:
        _print(f"      ✓ 全部通过 —— 转换无损。{report.summary}")
    else:
        _print(f"      ✗ 存在失败项 —— {report.summary}")
        for check in report.failed:
            _print(f"        - {check.name}: {check.detail}")

    if args.keep:
        _print(f"\n测试文件保留在: {out_dir}")
    return 0 if report.ok else 3


# --------------------------------------------------------------------------- #
# 入口
# --------------------------------------------------------------------------- #
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docm2docx",
        description="DOCM 批量无损转换为 DOCX(带 GUI / 命令行两种用法)",
    )
    sub = parser.add_subparsers(dest="command")

    p_scan = sub.add_parser("scan", help="扫描路径中的 .docm 文件")
    p_scan.add_argument("paths", nargs="+", help="分区根目录 / 文件夹 / 单文件")
    p_scan.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    p_scan.add_argument("--include-system", action="store_true", help="不跳过系统目录")
    p_scan.add_argument("--kernel", choices=[KERNEL_AUTO, KERNEL_EVERYTHING,
                                            KERNEL_WINDOWS_INDEX, KERNEL_WALK],
                        default=KERNEL_AUTO,
                        help="搜索内核:auto=自动(默认) everything=Everything 索引 "
                             "windowsindex=Windows 搜索索引 walk=直接遍历")
    p_scan.set_defaults(func=cmd_scan)

    p_conv = sub.add_parser("convert", help="批量转换 .docm → .docx")
    p_conv.add_argument("paths", nargs="+", help="分区根目录 / 文件夹 / 单文件")
    p_conv.add_argument("--out", default="", help="输出根目录(默认与源文件同目录)")
    p_conv.add_argument("--overwrite", choices=["skip", "overwrite", "rename"],
                        default="skip", help="目标已存在时的策略")
    p_conv.add_argument("--delete-source", choices=["none", "recycle", "delete"],
                        default="none",
                        help="转换成功后如何处理源 .docm:none=保留(默认) "
                             "recycle=移到回收站(可恢复) delete=永久删除")
    p_conv.add_argument("-y", "--yes", action="store_true",
                        help="跳过确认(配合 --delete-source 使用)")
    p_conv.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    p_conv.add_argument("--include-system", action="store_true", help="不跳过系统目录")
    p_conv.add_argument("--kernel", choices=[KERNEL_AUTO, KERNEL_EVERYTHING,
                                            KERNEL_WINDOWS_INDEX, KERNEL_WALK],
                        default=KERNEL_AUTO,
                        help="搜索内核:auto=自动(默认) everything=Everything 索引 "
                             "windowsindex=Windows 搜索索引 walk=直接遍历")
    p_conv.set_defaults(func=cmd_convert)

    p_test = sub.add_parser("selftest", help="生成测试 docm → 转换 → 机器校验")
    p_test.add_argument("--dir", default=os.path.join(os.getcwd(), "selftest_output"),
                        help="测试输出目录")
    p_test.add_argument("--keep", action="store_true", help="保留测试文件")
    p_test.set_defaults(func=cmd_selftest)
    return parser


def main(argv: Sequence[str] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    if not getattr(args, "func", None):
        parser.print_help()
        return 0
    return args.func(args)
