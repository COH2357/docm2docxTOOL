# -*- coding: utf-8 -*-
"""打发布包:便携版(含内置运行时)+ 纯源码版。

用法:
    python tools/package_release.py [版本号]      # 默认 1.0.0

产物输出到仓库的上一级目录(与仓库同级),避免把 zip 混进仓库:
    docm2docx-v<版本>-portable.zip   含 runtime/(Python + pywin32 + Everything)
    docm2docx-v<版本>-source.zip     纯源码,需自备 Python 3.8+

先决条件:便携版需要先跑 build-runtime.bat 生成 runtime/ 目录。
"""
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.dirname(ROOT)
VERSION = sys.argv[1] if len(sys.argv) > 1 else "1.0.0"
BASE = "docm2docx"

SKIP_DIRS = {".git", "__pycache__", "selftest_output", "_flow_ref", "tools"}
SKIP_FILES = {"Flow.Launcher-dev.zip", "启动错误.log"}


def collect(include_runtime: bool):
    items = []
    for dirpath, dirnames, filenames in os.walk(ROOT):
        rel_dir = os.path.relpath(dirpath, ROOT)
        parts = rel_dir.split(os.sep)
        if any(part in SKIP_DIRS for part in parts):
            continue
        if not include_runtime and parts and parts[0] == "runtime":
            continue
        for name in filenames:
            if name in SKIP_FILES or name.endswith(".pyc") or name.startswith("~$"):
                continue
            if name.startswith("启动错误"):
                continue
            full = os.path.join(dirpath, name)
            rel = os.path.join(BASE, os.path.relpath(full, ROOT))
            items.append((full, rel))
    return items


def build(items, out_path: str):
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for index, (full, rel) in enumerate(items, 1):
            zf.write(full, rel)
            if index % 500 == 0:
                print(f"    {index}/{len(items)}")
    size = os.path.getsize(out_path)
    print(f"  -> {os.path.basename(out_path)}  {size / 1024 / 1024:.1f} MB")
    return size


if __name__ == "__main__":
    print(f"打包 v{VERSION} …")
    print("便携版(含内置运行时)…")
    build(collect(True), os.path.join(OUT_DIR, f"{BASE}-v{VERSION}-portable.zip"))
    print("源码版(不含运行时)…")
    build(collect(False), os.path.join(OUT_DIR, f"{BASE}-v{VERSION}-source.zip"))
    print("完成。")
