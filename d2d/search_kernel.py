# -*- coding: utf-8 -*-
"""高速搜索内核。

移植自 Flow Launcher 的 Flow.Launcher.Plugin.Explorer/Search 思路,
按"快 → 慢"提供三档内核,并统一成同一个调用接口:

    1. EverythingKernel   —— 走 Everything 的 IPC 索引(毫秒级,覆盖全盘 NTFS 卷)
    2. WindowsIndexKernel —— 走 Windows 搜索索引(OLEDB Search.CollatorDSO)
    3. WalkKernel         —— 直接遍历文件系统(永远可用,但慢)

Everything 的 IPC 用的是 Flow Launcher 附带的 Everything SDK 助手库
(d2d/native/Everything.dll),通过 ctypes 调用,和官方 SDK 用法一致。
"""
from __future__ import annotations

import os
import time
from dataclasses import dataclass, field
from typing import Callable, List, Optional, Sequence, Tuple

from . import scanner
from .constants import DOCM_SUFFIX, TEMP_FILE_PREFIX

KERNEL_AUTO = "auto"
KERNEL_EVERYTHING = "everything"
KERNEL_WINDOWS_INDEX = "windowsindex"
KERNEL_WALK = "walk"

KERNEL_LABELS = {
    KERNEL_AUTO: "自动(Everything + Windows 索引,取并集)",
    KERNEL_EVERYTHING: "Everything 索引(极速,毫秒级)",
    KERNEL_WINDOWS_INDEX: "Windows 搜索索引(快)",
    KERNEL_WALK: "直接遍历(兼容性最好,较慢)",
}
KERNEL_BY_LABEL = {label: key for key, label in KERNEL_LABELS.items()}

# 索引里不该出现在候选清单中的路径片段
EXCLUDE_PATH_MARKERS = (
    "\\$RECYCLE.BIN\\",
    "\\SYSTEM VOLUME INFORMATION\\",
    "\\RECYCLER\\",
    "\\WINDOWS\\PREFETCH\\",
)

CancelCb = Callable[[], bool]
ProgressCb = Callable[[str, int], None]


@dataclass
class ScanResult:
    files: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    kernel: str = KERNEL_WALK
    kernel_name: str = "遍历"
    elapsed: float = 0.0
    stale_skipped: int = 0


# --------------------------------------------------------------------------- #
# 公共过滤
# --------------------------------------------------------------------------- #
def _is_excluded(path: str) -> bool:
    upper = path.upper()
    if any(marker in upper for marker in EXCLUDE_PATH_MARKERS):
        return True
    base = os.path.basename(path)
    if base.startswith(TEMP_FILE_PREFIX):
        return True
    return not base.lower().endswith(DOCM_SUFFIX)


def _finalize(paths: Sequence[str], verify: bool = True) -> Tuple[List[str], int]:
    """去重、过滤、剔除索引里的失效条目。"""
    result: List[str] = []
    seen = set()
    stale = 0
    for path in paths:
        if not path or _is_excluded(path):
            continue
        key = os.path.normcase(os.path.abspath(path))
        if key in seen:
            continue
        seen.add(key)
        if verify and not os.path.exists(path):
            stale += 1
            continue
        result.append(os.path.abspath(path))
    result.sort(key=lambda p: p.lower())
    return result, stale


# --------------------------------------------------------------------------- #
# 内核 1:Everything
# --------------------------------------------------------------------------- #
class EverythingKernel:
    """通过 Everything 的 IPC 索引检索(需要 Everything 正在运行)。

    若本机没有运行中的 Everything,会自动拉起 runtime/everything 里内置的那份,
    让整个工具在没装 Everything 的电脑上也能用上极速内核。
    """

    name = "Everything"
    key = KERNEL_EVERYTHING

    EVERYTHING_REQUEST_FULL_PATH = 0x0004
    EVERYTHING_MAX = 0xFFFFFFFF
    EVERYTHING_ERROR_IPC = 1
    EVERYTHING_WINDOW_CLASS = "EVERYTHING_TASKBAR_NOTIFICATION"

    def __init__(self, dll_path: Optional[str] = None,
                 everything_exe: Optional[str] = None):
        here = os.path.dirname(os.path.abspath(__file__))
        self.dll_path = dll_path or os.path.join(here, "native", "Everything.dll")
        self.everything_exe = everything_exe or os.path.join(
            os.path.dirname(here), "runtime", "everything", "Everything.exe")
        self._dll = None
        self._ctypes = None
        self._load_error: Optional[str] = None
        self._started_process = None
        self.started_bundled = False

    # -- 加载 -------------------------------------------------------------- #
    def _load(self):
        if self._dll is not None:
            return self._dll
        if self._load_error:
            raise RuntimeError(self._load_error)
        if not os.path.isfile(self.dll_path):
            self._load_error = f"找不到 Everything SDK:{self.dll_path}"
            raise RuntimeError(self._load_error)
        try:
            import ctypes
            dll = ctypes.WinDLL(self.dll_path)
            self._ctypes = ctypes
            dll.Everything_SetSearchW.argtypes = [ctypes.c_wchar_p]
            dll.Everything_SetSearchW.restype = ctypes.c_int
            dll.Everything_QueryW.argtypes = [ctypes.c_bool]
            dll.Everything_QueryW.restype = ctypes.c_bool
            dll.Everything_GetNumResults.restype = ctypes.c_int
            dll.Everything_GetTotResults.restype = ctypes.c_int
            dll.Everything_GetLastError.restype = ctypes.c_int
            dll.Everything_GetMajorVersion.restype = ctypes.c_int
            dll.Everything_SetMax.argtypes = [ctypes.c_int]
            dll.Everything_SetOffset.argtypes = [ctypes.c_int]
            dll.Everything_SetRequestFlags.argtypes = [ctypes.c_uint]
            dll.Everything_SetMatchPath.argtypes = [ctypes.c_bool]
            dll.Everything_GetResultFullPathNameW.argtypes = [
                ctypes.c_int, ctypes.c_wchar_p, ctypes.c_int]
            self._dll = dll
        except Exception as exc:
            self._load_error = f"加载 Everything SDK 失败:{exc}"
            raise RuntimeError(self._load_error) from exc
        return self._dll

    # -- 可用性 ------------------------------------------------------------ #
    def _ipc_ok(self) -> bool:
        """做一次真实查询来判断 IPC 是否连通(而不是只看错误码)。"""
        try:
            dll = self._load()
            dll.Everything_SetSearchW("*")
            ok = dll.Everything_QueryW(True)
            return bool(ok) and dll.Everything_GetLastError() != self.EVERYTHING_ERROR_IPC
        except Exception:
            return False

    def _everything_instance_exists(self) -> bool:
        """系统里是否已有 Everything 实例(通过它的 IPC 窗口类判断)。"""
        try:
            import ctypes
            from ctypes import wintypes

            found = []

            @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
            def callback(hwnd, _lparam):
                buffer = ctypes.create_unicode_buffer(256)
                ctypes.windll.user32.GetClassNameW(hwnd, buffer, 256)
                if buffer.value == self.EVERYTHING_WINDOW_CLASS:
                    found.append(True)
                    return False
                return True

            ctypes.windll.user32.EnumWindows(callback, 0)
            return bool(found)
        except Exception:
            return False

    def _start_bundled_everything(self, timeout: float = 20.0) -> bool:
        """拉起内置的 Everything,并等它的 IPC 窗口就绪。"""
        if not self.everything_exe or not os.path.isfile(self.everything_exe):
            return False
        try:
            import subprocess
            self._started_process = subprocess.Popen(
                [self.everything_exe, "-startup"],
                cwd=os.path.dirname(self.everything_exe))
        except Exception:
            return False
        deadline = time.time() + timeout
        while time.time() < deadline:
            time.sleep(0.5)
            if self._ipc_ok():
                # 已有别的实例时,内置副本会立即退出,这种就不算"我们启动了"
                self.started_bundled = self._started_process.poll() is None
                return True
        return False

    def available(self) -> bool:
        try:
            self._load()
        except Exception:
            return False
        if self._ipc_ok():
            return True
        # 已经有一个 Everything 实例了,却连不上 IPC —— 不再启动第二个,避免抢窗口
        if self._everything_instance_exists():
            return False
        return self._start_bundled_everything()

    def shutdown(self) -> None:
        """如果是我们启动的内置 Everything,退出时关掉它。"""
        process = self._started_process
        self._started_process = None
        if process is None:
            return
        try:
            process.terminate()
        except Exception:
            pass

    def status(self) -> str:
        try:
            dll = self._load()
            version = dll.Everything_GetMajorVersion()
            return f"Everything SDK 已加载(主版本 {version})"
        except Exception as exc:
            return str(exc)

    # -- 检索 -------------------------------------------------------------- #
    def _query(self, dll, search: str) -> List[str]:
        dll.Everything_SetOffset(0)
        dll.Everything_SetMax(self.EVERYTHING_MAX)
        dll.Everything_SetRequestFlags(self.EVERYTHING_REQUEST_FULL_PATH)
        dll.Everything_SetMatchPath(False)
        dll.Everything_SetSearchW(search)
        if not dll.Everything_QueryW(True):
            raise RuntimeError(f"Everything 查询失败(错误码 {dll.Everything_GetLastError()})")
        count = dll.Everything_GetNumResults()
        paths: List[str] = []
        buffer = self._ctypes.create_unicode_buffer(4096)
        for index in range(count):
            dll.Everything_GetResultFullPathNameW(index, buffer, 4096)
            if buffer.value:
                paths.append(buffer.value)
        return paths

    def find(self, roots: Sequence[str], recursive: bool,
             cancel: Optional[CancelCb] = None,
             progress: Optional[ProgressCb] = None) -> ScanResult:
        result = ScanResult(kernel=self.key, kernel_name=self.name)
        started = time.time()
        dll = self._load()
        collected: List[str] = []

        if not roots:
            # 全盘
            collected = self._query(dll, f"*{DOCM_SUFFIX}")
            if progress:
                progress("Everything 索引", len(collected))
        else:
            for root in roots:
                if cancel is not None and cancel():
                    break
                root = os.path.abspath(root)
                search = f"*{DOCM_SUFFIX} \"{root}\""
                try:
                    collected.extend(self._query(dll, search))
                except RuntimeError as exc:
                    result.errors.append(f"{root}:{exc}")
                    continue
                if progress:
                    progress(root, len(collected))

        files, stale = _finalize(collected)
        if not recursive and roots:
            files = [p for p in files
                     if any(os.path.dirname(p).lower() == os.path.normpath(r).lower()
                            for r in roots)]
        result.files = files
        result.stale_skipped = stale
        result.elapsed = time.time() - started
        return result


# --------------------------------------------------------------------------- #
# 内核 2:Windows 搜索索引
# --------------------------------------------------------------------------- #
class WindowsIndexKernel:
    """走 Windows 内置搜索索引(OLEDB Search.CollatorDSO)。"""

    name = "Windows 索引"
    key = KERNEL_WINDOWS_INDEX

    CONNECTION_STRING = ("Provider=Search.CollatorDSO;"
                         "Extended Properties='Application=Windows';")

    def __init__(self):
        self._connection = None
        self._error: Optional[str] = None

    def _connect(self):
        if self._connection is not None:
            return self._connection
        if self._error:
            raise RuntimeError(self._error)
        try:
            import win32com.client
            connection = win32com.client.Dispatch("ADODB.Connection")
            connection.Open(self.CONNECTION_STRING)
            self._connection = connection
        except Exception as exc:
            self._error = f"连接 Windows 搜索索引失败:{exc}"
            raise RuntimeError(self._error) from exc
        return self._connection

    def available(self) -> bool:
        try:
            self._connect()
            return True
        except Exception:
            return False

    def status(self) -> str:
        try:
            self._connect()
            return "Windows 搜索索引已连接"
        except Exception as exc:
            return str(exc)

    def _query(self, root: Optional[str]) -> List[str]:
        import win32com.client
        from urllib.parse import urlparse
        from urllib.request import url2pathname

        connection = self._connect()
        if root:
            scope = f"scope='file:{root}'"
            sql = (f"SELECT System.ItemUrl FROM SystemIndex "
                   f"WHERE {scope} AND System.FileName LIKE '%{DOCM_SUFFIX}'")
        else:
            # 全盘:不能加 scope 限定,否则一条都查不到
            sql = (f"SELECT System.ItemUrl FROM SystemIndex "
                   f"WHERE System.FileName LIKE '%{DOCM_SUFFIX}'")
        recordset = win32com.client.Dispatch("ADODB.Recordset")
        recordset.Open(sql, connection, 3, 1)      # adOpenStatic / adLockReadOnly
        paths: List[str] = []
        try:
            while not recordset.EOF:
                url = recordset.Fields.Item(0).Value
                if url:
                    paths.append(url2pathname(urlparse(url).path))
                recordset.MoveNext()
        finally:
            try:
                recordset.Close()
            except Exception:
                pass
        return paths

    def find(self, roots: Sequence[str], recursive: bool,
             cancel: Optional[CancelCb] = None,
             progress: Optional[ProgressCb] = None) -> ScanResult:
        result = ScanResult(kernel=self.key, kernel_name=self.name)
        started = time.time()
        collected: List[str] = []

        for root in (roots or [None]):
            if cancel is not None and cancel():
                break
            try:
                collected.extend(self._query(root))
            except Exception as exc:
                result.errors.append(f"{root or '全盘'}:{exc}")
                continue
            if progress:
                progress(root or "Windows 索引", len(collected))

        files, stale = _finalize(collected)
        result.files = files
        result.stale_skipped = stale
        result.elapsed = time.time() - started
        return result


# --------------------------------------------------------------------------- #
# 内核 3:遍历
# --------------------------------------------------------------------------- #
class WalkKernel:
    """直接遍历文件系统(兜底,永远可用)。"""

    name = "遍历"
    key = KERNEL_WALK

    def available(self) -> bool:
        return True

    def status(self) -> str:
        return "直接遍历文件系统"

    def find(self, roots: Sequence[str], recursive: bool,
             cancel: Optional[CancelCb] = None,
             progress: Optional[ProgressCb] = None) -> ScanResult:
        result = ScanResult(kernel=self.key, kernel_name=self.name)
        started = time.time()
        collected: List[str] = []
        for root in roots:
            if cancel is not None and cancel():
                break
            files, errors, _ = scanner.iter_docm_files(
                root, recursive=recursive, skip_system=True,
                cancel=cancel, progress=progress)
            collected.extend(files)
            result.errors.extend(errors)
        files, stale = _finalize(collected, verify=False)
        result.files = files
        result.stale_skipped = stale
        result.elapsed = time.time() - started
        return result


# --------------------------------------------------------------------------- #
# 编排
# --------------------------------------------------------------------------- #
_everything_singleton: Optional[EverythingKernel] = None


def get_everything_kernel() -> EverythingKernel:
    """Everything 内核单例 —— 记录我们是否拉起过内置实例,便于退出时清理。"""
    global _everything_singleton
    if _everything_singleton is None:
        _everything_singleton = EverythingKernel()
    return _everything_singleton


def shutdown_everything() -> None:
    """退出时关掉由本工具启动的内置 Everything(不影响用户自己开的)。"""
    global _everything_singleton
    if _everything_singleton is not None:
        _everything_singleton.shutdown()
        _everything_singleton = None


def build_kernels() -> List:
    return [get_everything_kernel(), WindowsIndexKernel(), WalkKernel()]


def find_docm_files(roots: Sequence[str],
                    recursive: bool = True,
                    kernel: str = KERNEL_AUTO,
                    cancel: Optional[CancelCb] = None,
                    progress: Optional[ProgressCb] = None,
                    skip_system: bool = True) -> ScanResult:
    """按指定内核查找 .docm。

    kernel=auto 时:同时跑 Everything 与 Windows 索引两个快内核并**取并集**
    (两者覆盖范围不同,合并才不会漏),只有当它们都没给出任何结果时,
    才降级到遍历磁盘 —— 否则"目标目录不在索引范围内"会被误判成没有文件。
    """
    roots = [os.path.abspath(r) for r in roots if r]

    if kernel == KERNEL_AUTO:
        collected: List[str] = []
        errors: List[str] = []
        used: List[str] = []
        for provider in (get_everything_kernel(), WindowsIndexKernel()):
            if cancel is not None and cancel():
                break
            try:
                if not provider.available():
                    errors.append(f"{provider.name}:不可用,已跳过")
                    continue
                outcome = provider.find(roots, recursive, cancel, progress)
                collected.extend(outcome.files)
                errors.extend(outcome.errors)
                used.append(provider.name)
            except Exception as exc:
                errors.append(f"{provider.name}: {exc}")

        if collected:
            files, stale = _finalize(collected)
            kernel_name = " + ".join(used)
            if get_everything_kernel().started_bundled:
                kernel_name += "(已自动启动内置 Everything)"
            return ScanResult(files=files, errors=errors,
                              kernel="auto", kernel_name=kernel_name,
                              elapsed=0.0, stale_skipped=stale)

        # 两个快内核都没结果,回退到遍历
        outcome = WalkKernel().find(roots, recursive, cancel, progress)
        outcome.errors = list(errors) + list(outcome.errors)
        outcome.kernel_name = "遍历(索引无结果)"
        return outcome

    if kernel == KERNEL_EVERYTHING:
        candidates = [get_everything_kernel()]
    elif kernel == KERNEL_WINDOWS_INDEX:
        candidates = [WindowsIndexKernel()]
    else:
        candidates = [WalkKernel()]

    errors: List[str] = []
    for provider in candidates:
        try:
            if not provider.available():
                errors.append(f"{provider.name}:不可用")
                if provider is candidates[-1]:
                    break
                continue
            outcome = provider.find(roots, recursive, cancel, progress)
            outcome.errors = list(outcome.errors) + errors
            return outcome
        except Exception as exc:
            errors.append(f"{provider.name}: {exc}")

    return ScanResult(files=[], errors=errors, kernel=kernel,
                      kernel_name=candidates[0].name)
