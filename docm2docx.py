#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""DOCM → DOCX 批量无损转换器(带图形界面)。

用法:
    python docm2docx.py              # 启动图形界面
    python docm2docx.py --check      # 环境自检(依赖 / Word 是否就绪)
    python docm2docx.py selftest     # 命令行自检:生成测试 docm → 转换 → 机器校验
    python docm2docx.py scan D:\\     # 扫描分区中的 .docm
    python docm2docx.py convert D:\\ --out E:\\out
"""
from __future__ import annotations

import os
import sys
import traceback

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

ERROR_LOG = os.path.join(APP_DIR, "启动错误.log")


def _report_crash(exc: BaseException) -> None:
    """启动失败时:落盘日志 + 弹窗,避免用 pythonw 启动时静默退出。"""
    detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    try:
        with open(ERROR_LOG, "w", encoding="utf-8") as handle:
            handle.write(detail)
    except OSError:
        pass
    try:
        sys.stderr.write(detail)
    except Exception:
        pass
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "启动失败",
            f"程序启动时出错:\n\n{exc}\n\n"
            f"详细信息已写入:\n{ERROR_LOG}\n\n"
            "常见原因:缺少 pywin32(请执行 pip install -r requirements.txt)"
            "、未安装 Microsoft Word、或 Word 被安全策略禁用。")
        root.destroy()
    except Exception:
        pass


def _check_dependencies() -> None:
    missing = []
    try:
        import pythoncom          # noqa: F401
        import win32com.client     # noqa: F401
    except ImportError:
        missing.append("pywin32")
    try:
        import tkinter             # noqa: F401
    except ImportError:
        missing.append("tkinter")
    if missing:
        raise RuntimeError(
            "缺少依赖: " + ", ".join(missing)
            + "\n请先执行:  pip install -r requirements.txt")


def _word_registered() -> bool:
    """检查 Windows 是否注册了 Word 的 COM ProgID(不启动 Word 进程)。"""
    try:
        import winreg
    except ImportError:
        return True
    for progid in ("Word.Application", "Word.Application.16", "Word.Application.15"):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
                return True
        except OSError:
            continue
    return False


def _run_check() -> int:
    """环境自检,输出给控制台看。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    ok = True
    print("=" * 52)
    print("DOCM → DOCX 转换器 · 环境自检")
    print("=" * 52)
    print(f"  Python          : {sys.version.split()[0]}  ({sys.executable})")

    try:
        import win32com.client      # noqa: F401
        print("  pywin32         : 已安装")
    except ImportError:
        print("  pywin32         : 未安装  →  pip install -r requirements.txt")
        ok = False

    try:
        import tkinter
        print(f"  tkinter         : 已安装 (Tk {tkinter.TkVersion})")
    except ImportError:
        print("  tkinter         : 未安装  →  重装 Python 时勾选 tcl/tk 组件")
        ok = False

    word = _word_registered()
    print(f"  Microsoft Word  : {'已就绪' if word else '未检测到 —— 请确认已安装 Word'}")
    if not word:
        ok = False

    # 便携运行时
    runtime_python = os.path.join(APP_DIR, "runtime", "python", "python.exe")
    runtime_everything = os.path.join(APP_DIR, "runtime", "everything", "Everything.exe")
    if os.path.isfile(runtime_python):
        print("  便携 Python     : 已内置(runtime\\python),启动时优先使用")
    else:
        print("  便携 Python     : 未内置(可运行 准备便携环境.bat 生成)")
    if os.path.isfile(runtime_everything):
        print("  便携 Everything : 已内置(runtime\\everything),需要时自动拉起")
    else:
        print("  便携 Everything : 未内置(可运行 准备便携环境.bat 生成)")

    print("=" * 52)
    print("  结论:" + ("环境正常,可以启动。" if ok
                       else "存在缺失项,请先按上面的提示处理。"))
    print("=" * 52)
    return 0 if ok else 1


def main() -> int:
    argv = sys.argv[1:]
    if argv and argv[0] == "--check":
        return _run_check()
    if argv:
        from d2d.cli import main as cli_main
        return cli_main(argv)

    _check_dependencies()
    from d2d.gui import main as gui_main
    return gui_main()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException as exc:                 # 兜底:不让窗口无声消失
        _report_crash(exc)
        sys.exit(1)
