# -*- coding: utf-8 -*-
"""扫描器:在分区 / 文件夹 / 单文件范围内定位所有 .docm。

设计要点:
  * 迭代式遍历(os.scandir + 显式栈),避免深目录触发递归深度限制;
  * 默认跳过回收站、系统目录、版本库等噪声目录,并可整体关闭"跳过系统目录";
  * 扫描过程可取消、可上报进度,适合在 GUI 中跑整盘。
"""
from __future__ import annotations

import ctypes
import os
import time
from dataclasses import dataclass
from typing import Callable, List, Optional, Sequence, Tuple

from .constants import DOCM_SUFFIX, TEMP_FILE_PREFIX

# 无论是否"跳过系统目录"都不进入的目录(垃圾目录 / 缓存 / 版本库)
UNCONDITIONAL_SKIP = {
    "$recycle.bin", "system volume information", "perflogs",
    "$windows.~bt", "$windows.~ws", "windows.old", "recovery",
    "node_modules", "__pycache__", ".git", ".hg", ".svn",
    ".venv", "venv", "site-packages", ".idea", ".vscode",
}

# 开启"跳过系统目录"时额外跳过的目录
SYSTEM_SKIP = {
    "windows", "program files", "program files (x86)", "program files (arm)",
    "programdata", "appdata", "$recycle.bin", "system volume information",
    "windows.old", "$windows.~bt", "$windows.~ws", "recovery", "perflogs",
    "oneDriveTemp", "$windows.~q",
}

# 仅在盘符根目录下才跳过的名字(避免误伤用户自建的同名目录)
ROOT_ONLY_SYSTEM_SKIP = {
    "windows", "program files", "program files (x86)", "program files (arm)",
}

ProgressCb = Callable[[str, int], None]     # (当前目录, 已找到文件数)
CancelCb = Callable[[], bool]


# --------------------------------------------------------------------------- #
# 磁盘枚举
# --------------------------------------------------------------------------- #
@dataclass
class Drive:
    letter: str
    label: str = ""
    fs_type: str = ""
    total: int = 0
    free: int = 0

    @property
    def path(self) -> str:
        return self.letter

    @property
    def display(self) -> str:
        text = f"{self.letter} {self.label}".strip()
        if self.fs_type:
            text += f" [{self.fs_type}]"
        if self.total:
            text += f" — 可用 {human_size(self.free)} / 共 {human_size(self.total)}"
        return text


def human_size(num: float) -> str:
    """字节数转可读字符串。"""
    if num is None:
        return "-"
    for unit in ("B", "KB", "MB", "GB", "TB", "PB"):
        if abs(num) < 1024.0:
            return f"{num:.1f} {unit}" if unit != "B" else f"{int(num)} {unit}"
        num /= 1024.0
    return f"{num:.1f} EB"


def list_drives() -> List[Drive]:
    """列出本机所有逻辑驱动器(含未挂载的光驱,类型会标注)。"""
    drives: List[Drive] = []
    try:
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    except Exception:
        return drives

    kernel32 = ctypes.windll.kernel32
    try:
        kernel32.GetVolumeInformationW.argtypes = [
            ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_ulong,
            ctypes.POINTER(ctypes.c_ulong), ctypes.POINTER(ctypes.c_ulong),
            ctypes.POINTER(ctypes.c_ulong), ctypes.c_wchar_p, ctypes.c_ulong,
        ]
        kernel32.GetVolumeInformationW.restype = ctypes.c_bool
    except Exception:
        pass

    for index in range(26):
        if not (bitmask & (1 << index)):
            continue
        letter = f"{chr(ord('A') + index)}:\\"
        label = ""
        fs_type = ""
        try:
            name_buf = ctypes.create_unicode_buffer(261)
            fs_buf = ctypes.create_unicode_buffer(261)
            serial = ctypes.c_ulong()
            max_comp = ctypes.c_ulong()
            flags = ctypes.c_ulong()
            ok = kernel32.GetVolumeInformationW(
                ctypes.c_wchar_p(letter), name_buf, 261,
                ctypes.byref(serial), ctypes.byref(max_comp), ctypes.byref(flags),
                fs_buf, 261,
            )
            if ok:
                label = name_buf.value
                fs_type = fs_buf.value
        except Exception:
            pass

        total = free = 0
        try:
            import shutil
            usage = shutil.disk_usage(letter)
            total, free = usage.total, usage.free
        except Exception:
            pass

        drives.append(Drive(letter=letter, label=label, fs_type=fs_type,
                            total=total, free=free))
    return drives


# --------------------------------------------------------------------------- #
# 目录遍历
# --------------------------------------------------------------------------- #
def _should_skip_dir(name: str, path: str, skip_system: bool) -> bool:
    low = name.lower()
    if low in UNCONDITIONAL_SKIP:
        return True
    if skip_system:
        if low in SYSTEM_SKIP:
            # windows / program files 只在盘根下跳过,更深层的同名目录保留
            if low in ROOT_ONLY_SYSTEM_SKIP:
                parent = os.path.dirname(os.path.normpath(path))
                if len(parent) <= 3:      # 形如 "C:\" —— 盘根下一级
                    return True
                return False
            return True
    return False


def iter_docm_files(root: str,
                    recursive: bool = True,
                    skip_system: bool = True,
                    cancel: Optional[CancelCb] = None,
                    progress: Optional[ProgressCb] = None,
                    interval: float = 0.25) -> Tuple[List[str], List[str], int]:
    """遍历 root,返回 (docm 文件列表, 访问错误列表, 已扫描目录数)。"""
    found: List[str] = []
    errors: List[str] = []
    scanned = 0

    root = os.path.abspath(root)
    if not os.path.isdir(root):
        return found, [f"目录不存在或不可访问: {root}"], 0

    stack = [root]
    last_report = 0.0
    while stack:
        if cancel is not None and cancel():
            break
        current = stack.pop()
        scanned += 1
        try:
            with os.scandir(current) as iterator:
                entries = list(iterator)
        except (PermissionError, OSError) as exc:
            errors.append(f"{current} — {getattr(exc, 'strerror', None) or exc}")
            continue

        subdirs: List[str] = []
        for entry in entries:
            try:
                if entry.is_dir(follow_symlinks=False):
                    if not recursive:
                        continue
                    if _should_skip_dir(entry.name, entry.path, skip_system):
                        continue
                    subdirs.append(entry.path)
                elif entry.is_file(follow_symlinks=False):
                    if entry.name.startswith(TEMP_FILE_PREFIX):
                        continue          # Word 正在打开的临时锁文件
                    if entry.name.lower().endswith(DOCM_SUFFIX):
                        found.append(entry.path)
            except OSError:
                continue
        stack.extend(subdirs)

        now = time.time()
        if progress is not None and (now - last_report) >= interval:
            last_report = now
            progress(current, len(found))

    if progress is not None:
        progress("", len(found))
    found.sort(key=lambda p: p.lower())
    return found, errors, scanned


def collect_docm(targets: Sequence[Tuple[str, str]],
                 recursive: bool = True,
                 skip_system: bool = True,
                 cancel: Optional[CancelCb] = None,
                 progress: Optional[ProgressCb] = None) -> Tuple[List[str], List[str], int]:
    """对多个目标(类型, 路径)统一扫描。

    目标类型: "drive" / "folder" / "file"。
    返回 (docm 文件列表, 错误列表, 已扫描目录数)。
    """
    all_files: List[str] = []
    errors: List[str] = []
    scanned_dirs = 0
    offset = 0

    for kind, path in targets:
        if cancel is not None and cancel():
            break
        if kind == "file":
            if path.lower().endswith(DOCM_SUFFIX) and os.path.isfile(path):
                if not os.path.basename(path).startswith(TEMP_FILE_PREFIX):
                    all_files.append(os.path.abspath(path))
            elif not path.lower().endswith(DOCM_SUFFIX):
                errors.append(f"跳过(非 .docm): {path}")
            continue

        def report(current_dir: str, found: int, _offset: int = offset) -> None:
            if progress is not None:
                progress(current_dir, _offset + found)

        files, errs, scanned = iter_docm_files(
            path, recursive=recursive, skip_system=skip_system,
            cancel=cancel, progress=report,
        )
        all_files.extend(files)
        errors.extend(errs)
        scanned_dirs += scanned
        offset += len(files)

    # 去重(不同目标可能覆盖到同一文件)
    unique: List[str] = []
    seen = set()
    for path in sorted(all_files, key=lambda p: p.lower()):
        key = os.path.normcase(os.path.abspath(path))
        if key in seen:
            continue
        seen.add(key)
        unique.append(path)
    return unique, errors, scanned_dirs


def file_info(path: str) -> dict:
    """取文件元信息(大小 / 修改时间),失败返回默认值。"""
    info = {"path": path, "size": 0, "mtime": ""}
    try:
        stat = os.stat(path)
        info["size"] = stat.st_size
        info["mtime"] = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(stat.st_mtime))
    except OSError:
        pass
    return info
