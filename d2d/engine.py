# -*- coding: utf-8 -*-
"""Word COM 转换引擎:docm → docx 无损另存。

关键点:
  * 用 DispatchEx 起独立的后台 Word 实例,不干扰用户正在打开的文档;
  * AutomationSecurity = ForceDisable,打开 .docm 时不执行任何宏;
  * ReadOnly 打开 + SaveAs2(FileFormat=16) 另存,原 .docm 保持不动;
  * 全程 DisplayAlerts=None,避免"宏将丢失"等弹窗卡住批处理。
"""
from __future__ import annotations

import os
import time
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Sequence, Tuple

import pythoncom
import win32com.client

from .constants import (
    CLEANUP_DELETE, CLEANUP_NONE, CLEANUP_RECYCLE, DOCX_SUFFIX,
    MSO_AUTOMATION_SECURITY_FORCE_DISABLE, OVERWRITE_RENAME,
    OVERWRITE_REPLACE, OVERWRITE_SKIP, WD_ALERTS_NONE, WD_DO_NOT_SAVE_CHANGES,
    WD_FORMAT_DOCUMENT_DEFAULT, WD_FORMAT_XML_DOCUMENT_MACRO,
)
from .repair import repair_document, replace_source

STATUS_PENDING = "待转换"
STATUS_SUCCESS = "成功"
STATUS_SKIPPED = "跳过"
STATUS_FAILED = "失败"


class WordUnavailable(RuntimeError):
    """Word 无法启动(未安装 / COM 注册缺失 / 被策略禁用)。"""


# --------------------------------------------------------------------------- #
# 进程辅助:确保后台 Word 实例不会残留
# --------------------------------------------------------------------------- #
def _winword_pids() -> set:
    """列出当前所有 WINWORD.EXE 的进程 ID。"""
    pids: set = set()
    try:
        import ctypes
        from ctypes import wintypes

        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [
                ("dwSize", wintypes.DWORD),
                ("cntUsage", wintypes.DWORD),
                ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
                ("th32ModuleID", wintypes.DWORD),
                ("cntThreads", wintypes.DWORD),
                ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", ctypes.c_long),
                ("dwFlags", wintypes.DWORD),
                ("szExeFile", ctypes.c_char * 260),
            ]

        kernel32 = ctypes.windll.kernel32
        kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
        snapshot = kernel32.CreateToolhelp32Snapshot(0x00000002, 0)  # SNAPPROCESS
        if snapshot == -1 or snapshot == 0xFFFFFFFF:
            return pids
        try:
            entry = PROCESSENTRY32()
            entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
            if kernel32.Process32First(snapshot, ctypes.byref(entry)):
                while True:
                    if entry.szExeFile.decode("ascii", "ignore").lower() == "winword.exe":
                        pids.add(int(entry.th32ProcessID))
                    if not kernel32.Process32Next(snapshot, ctypes.byref(entry)):
                        break
        finally:
            kernel32.CloseHandle(snapshot)
    except Exception:
        pass
    return pids


def _terminate(pid: int) -> bool:
    """强制结束指定进程。"""
    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32
        kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel32.OpenProcess.restype = wintypes.HANDLE
        handle = kernel32.OpenProcess(0x0001, False, pid)   # PROCESS_TERMINATE
        if not handle:
            return False
        try:
            kernel32.TerminateProcess(handle, 0)
        finally:
            kernel32.CloseHandle(handle)
        return True
    except Exception:
        return False


def _process_command_line(pid: int) -> str:
    """读取进程命令行(用于区分自动化实例与用户手动打开的 Word)。"""
    try:
        import subprocess
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f"(Get-CimInstance Win32_Process -Filter 'ProcessId={pid}').CommandLine"],
            capture_output=True, text=True, timeout=25,
        )
        return (result.stdout or "").strip()
    except Exception:
        return ""


# --------------------------------------------------------------------------- #
# 数据结构
# --------------------------------------------------------------------------- #
@dataclass
class ConvertItem:
    src: str
    dst: str = ""
    root: str = ""                       # 扫描时的根目录,用于计算相对输出路径
    status: str = STATUS_PENDING
    message: str = ""
    elapsed: float = 0.0
    repaired: bool = False               # 是否走了“修复损坏文档”流程

    @property
    def name(self) -> str:
        return os.path.basename(self.src)


@dataclass
class ConvertSummary:
    total: int = 0
    success: int = 0
    skipped: int = 0
    failed: int = 0
    removed: int = 0             # 转换成功后被处理掉的源 .docm 数量
    cleanup_failed: int = 0      # 源文件处理失败的数量
    repaired: int = 0            # 先修复损坏文档、再转换成功的数量
    elapsed: float = 0.0
    cancelled: bool = False
    errors: List[Tuple[str, str]] = field(default_factory=list)

    @property
    def done(self) -> int:
        return self.success + self.skipped + self.failed


ItemCb = Callable[[ConvertItem, int, int], None]     # (item, 已完成, 总数)
LogCb = Callable[[str], None]


# --------------------------------------------------------------------------- #
# 工具函数
# --------------------------------------------------------------------------- #
def unique_path(path: str) -> str:
    """生成不冲突的文件名: name.docx → name (1).docx → name (2).docx …"""
    base, ext = os.path.splitext(path)
    index = 1
    while os.path.exists(f"{base} ({index}){ext}"):
        index += 1
    return f"{base} ({index}){ext}"


def unblock_file(path: str) -> bool:
    """移除来自互联网的阻止标记(MOTW),避免受保护视图导致转换失败。

    只删除 NTFS 备用数据流 Zone.Identifier,不触碰文档内容本身。
    """
    ads = path + ":Zone.Identifier"
    try:
        with open(ads, "rb"):
            pass
    except OSError:
        return False
    try:
        os.remove(ads)
        return True
    except OSError:
        return False


def plan_output(src: str, root: str = "", output_root: str = "") -> str:
    """计算 docm → docx 的目标路径。"""
    src = os.path.abspath(src)
    if output_root:
        relative = os.path.basename(src)
        if root:
            try:
                relative = os.path.relpath(src, root)
            except ValueError:
                relative = os.path.basename(src)
        target = os.path.join(os.path.abspath(output_root),
                              os.path.splitext(relative)[0] + DOCX_SUFFIX)
        return os.path.abspath(target)
    return os.path.splitext(src)[0] + DOCX_SUFFIX


# --------------------------------------------------------------------------- #
# 源文件删除
# --------------------------------------------------------------------------- #
def send_to_recycle_bin(paths) -> Tuple[bool, str]:
    """把文件移入 Windows 回收站(可恢复)。返回 (是否成功, 说明)。"""
    paths = [os.path.abspath(p) for p in (paths if isinstance(paths, (list, tuple)) else [paths])]
    paths = [p for p in paths if os.path.exists(p)]
    if not paths:
        return True, ""
    try:
        import ctypes
        from ctypes import wintypes

        class SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [
                ("hwnd", wintypes.HWND),
                ("wFunc", wintypes.UINT),
                ("pFrom", wintypes.LPCWSTR),
                ("pTo", wintypes.LPCWSTR),
                ("fFlags", wintypes.WORD),
                ("fAnyOperationsAborted", wintypes.BOOL),
                ("hNameMappings", ctypes.c_void_p),
                ("lpszProgressTitle", wintypes.LPCWSTR),
            ]

        FO_DELETE = 0x0003
        FOF_SILENT = 0x0004
        FOF_NOCONFIRMATION = 0x0010
        FOF_ALLOWUNDO = 0x0040        # 关键:允许撤销 = 进回收站而不是真删
        FOF_NOERRORUI = 0x0400

        buffer = ctypes.create_unicode_buffer("\0".join(paths) + "\0\0")
        op = SHFILEOPSTRUCTW()
        op.hwnd = None
        op.wFunc = FO_DELETE
        op.pFrom = ctypes.cast(buffer, wintypes.LPCWSTR)
        op.pTo = None
        op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
        op.fAnyOperationsAborted = False
        op.hNameMappings = None
        op.lpszProgressTitle = None

        code = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
        if code != 0:
            return False, f"回收站操作失败(错误码 {code})"
        if op.fAnyOperationsAborted:
            return False, "回收站操作被中止"
        return True, ""
    except Exception as exc:
        return False, f"回收站操作异常: {exc}"


def remove_source(path: str, mode: str) -> Tuple[bool, str]:
    """按策略处理转换完成后的源 .docm。返回 (是否成功, 说明)。"""
    if mode == CLEANUP_NONE:
        return True, ""
    if mode == CLEANUP_RECYCLE:
        ok, message = send_to_recycle_bin(path)
        return ok, ("已移入回收站" if ok else message)
    if mode == CLEANUP_DELETE:
        try:
            import stat
            os.chmod(path, stat.S_IWRITE)      # 解除只读属性
            os.remove(path)
            return True, "已永久删除"
        except OSError as exc:
            return False, f"删除失败: {exc}"
    return True, ""


# --------------------------------------------------------------------------- #
# Word 会话
# --------------------------------------------------------------------------- #
class WordSession:
    """管理一个后台 Word 实例的生命周期(必须在同一线程内创建/销毁)。"""

    def __init__(self, visible: bool = False):
        self.visible = visible
        self.app = None
        self._pids_before: set = set()
        self._owned_pids: set = set()

    # -- 上下文管理 -------------------------------------------------------- #
    def __enter__(self) -> "WordSession":
        self.start()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    # -- 生命周期 ---------------------------------------------------------- #
    def start(self) -> None:
        self._pids_before = _winword_pids()
        pythoncom.CoInitialize()
        try:
            self.app = win32com.client.DispatchEx("Word.Application")
        except Exception as exc:                     # pragma: no cover - 环境相关
            pythoncom.CoUninitialize()
            raise WordUnavailable(
                "无法启动 Microsoft Word,请确认已安装 Word 且未被安全策略禁用。"
                f"\n原始错误: {exc}"
            ) from exc

        app = self.app
        app.Visible = self.visible
        app.DisplayAlerts = WD_ALERTS_NONE
        for setter, value in (
            ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
            ("ConfirmConversions", False),
            ("UpdateLinksAtOpen", False),
            ("WarnBeforeSavingPrintingSendingMarkup", False),
            ("DoNotPromptForConvert", True),
            ("SaveInterval", 0),
            ("BackgroundSave", False),
        ):
            try:
                if setter == "AutomationSecurity":
                    setattr(app, setter, value)
                else:
                    setattr(app.Options, setter, value)
            except Exception:
                pass      # 个别版本不支持某选项,忽略即可

        # 记录本次新建的 WINWORD 进程,便于退出时确认已彻底结束
        time.sleep(0.2)
        self._owned_pids = _winword_pids() - self._pids_before

    def close(self) -> None:
        app, self.app = self.app, None
        if app is not None:
            try:
                while app.Documents.Count > 0:
                    app.Documents(1).Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
            try:
                app.Quit(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        del app
        self._reap()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def _reap(self, timeout: float = 8.0) -> None:
        """等待后台 Word 进程退出;顽固不退出的直接结束,避免残留占用内存。

        只结束本次会话新建的进程,并且要求它确实是 COM 自动化实例
        (命令行含 /Automation 或 -Embedding),绝不碰用户手动打开的 Word。
        """
        if not self._owned_pids:
            return
        deadline = time.time() + timeout
        alive = set(self._owned_pids)
        while time.time() < deadline:
            alive = self._owned_pids & _winword_pids()
            if not alive:
                return
            time.sleep(0.25)
        for pid in sorted(alive):
            cmdline = _process_command_line(pid)
            if cmdline and ("/Automation" not in cmdline
                            and "-Embedding" not in cmdline):
                continue          # 不是自动化实例,放过(理论上不会发生)
            _terminate(pid)
        time.sleep(0.3)


# --------------------------------------------------------------------------- #
# 单文件转换
# --------------------------------------------------------------------------- #
@dataclass
class ConversionResult:
    """一次单文件转换的结果。"""

    status: str
    message: str
    repaired: bool = False


def _prepare_destination(dst: str, overwrite: str) -> Tuple[str, str, str]:
    """处理“目标已存在”的情况。返回 (dst, 状态, 说明);状态为空表示继续。"""
    if not os.path.exists(dst):
        return dst, "", ""
    if overwrite == OVERWRITE_SKIP:
        return "", STATUS_SKIPPED, "目标文件已存在"
    if overwrite == OVERWRITE_RENAME:
        return unique_path(dst), "", ""
    try:
        os.remove(dst)
    except OSError as exc:
        return "", STATUS_FAILED, f"无法覆盖已有文件: {exc}"
    return dst, "", ""


def _save_docx(word_app, src: str, dst: str) -> Tuple[str, str]:
    """用 Word 把 src 另存为 .docx。返回 (状态, 说明)。"""
    doc = None
    try:
        doc = word_app.Documents.Open(
            FileName=src,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
        saver = getattr(doc, "SaveAs2", None) or doc.SaveAs
        saver(FileName=dst, FileFormat=WD_FORMAT_DOCUMENT_DEFAULT,
              AddToRecentFiles=False)
        if not os.path.isfile(dst) or os.path.getsize(dst) == 0:
            return STATUS_FAILED, "另存后文件不存在或为空"
        return STATUS_SUCCESS, f"{os.path.getsize(dst):,} 字节"
    except Exception as exc:
        return STATUS_FAILED, str(exc).replace("\r", " ").replace("\n", " ")
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass


def _discard(path: str) -> None:
    try:
        if path and os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass


def convert_one(word_app, src: str, dst: str, overwrite: str = OVERWRITE_SKIP,
                repair: bool = False,
                log: Optional[Callable[[str], None]] = None) -> ConversionResult:
    """把单个 .docm 转成 .docx。

    ``repair=True`` 时,若正常转换失败,会尝试修复损坏的 .docm:
    先用 Word 自带的“打开并修复”,失败再修 OPC 包结构,然后由 Word
    另存为 .docm。修复成功后会**覆盖原 .docm**(同时去掉宏工程),再继续
    转成 .docx。
    """
    src = os.path.abspath(src)
    dst = os.path.abspath(dst)

    if not os.path.isfile(src):
        return ConversionResult(STATUS_FAILED, "源文件不存在")
    if os.path.getsize(src) == 0:
        return ConversionResult(STATUS_FAILED, "源文件为空")

    dst, prep_status, prep_message = _prepare_destination(dst, overwrite)
    if prep_status:
        return ConversionResult(prep_status, prep_message)

    # Word 另存到不存在的目录会报"这不是有效文件名",这里先确保目录存在
    parent = os.path.dirname(dst)
    if parent and not os.path.isdir(parent):
        try:
            os.makedirs(parent, exist_ok=True)
        except OSError as exc:
            return ConversionResult(STATUS_FAILED, f"无法创建输出目录: {exc}")

    status, message = _save_docx(word_app, src, dst)
    if status == STATUS_SUCCESS or not repair:
        return ConversionResult(status, message)

    # ---- 转换失败:尝试修复损坏的文档 ------------------------------- #
    first_error = message
    if log:
        log(f"  转换失败,尝试修复: {os.path.basename(src)} — {first_error}")

    outcome = repair_document(word_app, src, log=log)
    if not outcome.ok:
        if log:
            log(f"  修复失败: {outcome.detail}")
        return ConversionResult(STATUS_FAILED, f"{first_error}｜修复失败: {outcome.detail}")

    # 按要求先把修复后的 .docm 覆盖回原文件;覆盖失败则退而用临时副本转换
    replaced, replace_note = replace_source(src, outcome.path)
    convert_src = src if replaced else outcome.path
    if log:
        if replaced:
            log("  已修复并覆盖原 .docm(宏工程已去除),继续转换 …")
        else:
            log(f"  已修复,但覆盖原文件失败({replace_note}),改用临时副本转换 …")

    dst, _, _ = _prepare_destination(dst, OVERWRITE_REPLACE)
    status2, message2 = _save_docx(word_app, convert_src, dst)
    if status2 == STATUS_SUCCESS:
        if replaced:
            detail = "已修复并覆盖原 .docm(已去宏)"
        else:
            detail = f"已修复为临时副本(覆盖原文件失败: {replace_note})"
        return ConversionResult(STATUS_SUCCESS, f"{message2} | {detail}", repaired=True)

    if not replaced:
        _discard(outcome.path)
    extra = "原 .docm 已修复为可打开的 .docm" if replaced else ""
    return ConversionResult(
        STATUS_FAILED,
        f"{first_error}｜修复后仍无法转换: {message2}" + (f"｜{extra}" if extra else ""))



# --------------------------------------------------------------------------- #
# 批量转换
# --------------------------------------------------------------------------- #
def convert_batch(items: Sequence[ConvertItem],
                  overwrite: str = OVERWRITE_SKIP,
                  output_root: str = "",
                  cancel: Optional[Callable[[], bool]] = None,
                  on_item: Optional[ItemCb] = None,
                  log: Optional[LogCb] = None,
                  unblock: bool = True,
                  cleanup: str = CLEANUP_NONE,
                  repair: bool = False,
                  session: Optional[WordSession] = None) -> ConvertSummary:
    """批量转换。item.dst 为空时按 output_root/源目录自动推导。

    cleanup 控制转换成功后如何处理源 .docm:
      CLEANUP_NONE    不动(默认)
      CLEANUP_RECYCLE 移到回收站(可恢复)
      CLEANUP_DELETE  永久删除
    只有转换**成功**的文件才会被处理,跳过/失败的一律保留。

    repair=True 时,转换失败会先尝试修复损坏的 .docm(见 convert_one)。
    修复成功会覆盖原 .docm,所以 cleanup 同样作用于修复后的文件。
    """
    summary = ConvertSummary(total=len(items))
    if not items:
        return summary

    started = time.time()
    own_session = session is None
    word_session = session or WordSession(visible=False)

    def emit(message: str) -> None:
        if log:
            log(message)

    try:
        if own_session:
            emit("正在启动 Microsoft Word 后台实例 …")
            word_session.start()
        word_app = word_session.app

        for index, item in enumerate(items, 1):
            if cancel is not None and cancel():
                summary.cancelled = True
                emit("已取消,停止转换。")
                break

            if not item.dst:
                item.dst = plan_output(item.src, item.root, output_root)

            if unblock and os.path.isfile(item.src):
                try:
                    if unblock_file(item.src):
                        emit(f"已移除阻止标记: {os.path.basename(item.src)}")
                except Exception:
                    pass

            item.elapsed = 0.0
            begin = time.time()
            result = convert_one(word_app, item.src, item.dst, overwrite,
                                 repair=repair, log=emit)
            item.elapsed = time.time() - begin
            item.status = result.status
            item.message = result.message
            item.repaired = result.repaired

            if result.status == STATUS_SUCCESS:
                summary.success += 1
                if result.repaired:
                    summary.repaired += 1
            elif result.status == STATUS_SKIPPED:
                summary.skipped += 1
            else:
                summary.failed += 1
                summary.errors.append((item.src, result.message))

            # 转换成功后才处理源文件;跳过/失败的一律保留
            if result.status == STATUS_SUCCESS and cleanup != CLEANUP_NONE:
                removed, note = remove_source(item.src, cleanup)
                if removed:
                    summary.removed += 1
                    item.message = f"{item.message} | 源文件{note}"
                else:
                    summary.cleanup_failed += 1
                    item.message = f"{item.message} | 源文件{note}"
                emit(f"  {os.path.basename(item.src)}: 源文件{note}")

            if on_item:
                on_item(item, index, len(items))
    except WordUnavailable as exc:
        summary.failed = summary.total
        summary.errors.append(("", str(exc)))
        emit(f"错误: {exc}")
    finally:
        if own_session:
            word_app = None          # 先释放 COM 引用,再关闭实例
            word_session.close()

    summary.elapsed = time.time() - started
    return summary


# --------------------------------------------------------------------------- #
# 另存为其它格式(供测试模块使用)
# --------------------------------------------------------------------------- #
def save_as(word_app, doc, path: str, file_format: int) -> None:
    """把当前文档另存为指定格式。"""
    saver = getattr(doc, "SaveAs2", None) or doc.SaveAs
    saver(FileName=os.path.abspath(path), FileFormat=file_format,
          AddToRecentFiles=False)


def create_docm(word_app, path: str) -> None:
    """新建空白文档并另存为 .docm(FileFormat=13)。"""
    doc = word_app.Documents.Add()
    try:
        save_as(word_app, doc, path, WD_FORMAT_XML_DOCUMENT_MACRO)
    finally:
        try:
            doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        except Exception:
            pass
