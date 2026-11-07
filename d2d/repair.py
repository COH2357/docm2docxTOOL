# -*- coding: utf-8 -*-
"""损坏 .docm 的修复:Word 自带修复 + OPC 包结构修复。

背景
----
被宏病毒感染/破坏的 .docm,常见损坏并不在文档正文,而在 OPC 包结构:

* ``_rels/.rels`` 等关系部件里出现**重复的关系 Id**(同一 rId 指向不同目标);
* 压缩包里出现**同名重复部件**(例如 customUI/customUI.xml 两份);
* 关系指向已被删除/从未存在的部件。

Word 直接打开这种文件会报“文件可能已经损坏”,但手工用 Word 的
“打开并修复”通常能救回来。麻烦的是 COM 的
``Documents.Open(OpenAndRepair=True)`` **不会真正触发**这类关系级修复
(实测仍报同样的错),所以这里分两步:

1. 先按用户要求尝试 Word 自带的“打开并修复”并另存为 .docm;
2. 不行就自己修 OPC 包结构(去重部件、修正关系 Id、移除失效关系),
   再交给 Word 另存为 .docm —— 最终仍由 Word 产出干净的 .docm。

修复产物会**去掉 VBA 宏工程**:本工具的用途就是消除宏病毒,修复后的
.docm 若保留病毒宏,一旦被保留下来反而危险。文档正文/样式/图片不受
影响。

所有操作只使用 zip/XML 读写与 Word COM;打开文档时调用方必须已经
设置 ``AutomationSecurity = ForceDisable``(WordSession 已处理)。
"""
from __future__ import annotations

import os
import posixpath
import shutil
import stat
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import List, Optional, Tuple
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from .constants import WD_FORMAT_XML_DOCUMENT_MACRO
from .ooxml import has_macros, has_part, validate_package

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# 需要剔除的宏相关部件(小写比较)
MACRO_PARTS = {"word/vbaproject.bin", "word/vbadata.xml"}

TEMP_PREFIX = "~d2d-repair-"


class RepairError(RuntimeError):
    """修复流程无法继续。"""


# --------------------------------------------------------------------------- #
# 修复结果
# --------------------------------------------------------------------------- #
@dataclass
class PackageReport:
    """OPC 包结构修复的明细。"""

    duplicate_parts: List[str] = field(default_factory=list)
    renamed_rel_ids: List[str] = field(default_factory=list)
    duplicate_rels: List[str] = field(default_factory=list)
    dangling_rels: List[str] = field(default_factory=list)
    removed_parts: List[str] = field(default_factory=list)
    removed_macro_rels: List[str] = field(default_factory=list)
    removed_overrides: List[str] = field(default_factory=list)

    @property
    def changed(self) -> bool:
        return bool(self.duplicate_parts or self.renamed_rel_ids
                    or self.duplicate_rels or self.dangling_rels
                    or self.removed_parts or self.removed_macro_rels
                    or self.removed_overrides)

    def summary(self) -> str:
        bits: List[str] = []
        if self.duplicate_parts:
            bits.append(f"去重部件 {len(self.duplicate_parts)} 个")
        if self.renamed_rel_ids:
            bits.append(f"修正重复关系 Id {len(self.renamed_rel_ids)} 处")
        if self.duplicate_rels:
            bits.append(f"删除重复关系 {len(self.duplicate_rels)} 处")
        if self.dangling_rels:
            bits.append(f"移除失效关系 {len(self.dangling_rels)} 处")
        if self.removed_macro_rels:
            bits.append(f"移除宏关系 {len(self.removed_macro_rels)} 处")
        if self.removed_parts:
            bits.append(f"移除宏部件 {len(self.removed_parts)} 个")
        if self.removed_overrides:
            bits.append(f"清理内容类型 {len(self.removed_overrides)} 项")
        return "、".join(bits) if bits else "无需改动"


@dataclass
class RepairResult:
    """一次修复尝试的结果。"""

    ok: bool
    path: str = ""          # 修复后的 .docm(临时路径,调用方负责覆盖/清理)
    method: str = ""        # word / package
    detail: str = ""
    notes: List[str] = field(default_factory=list)
    in_place: bool = True   # 临时文件是否与源文件在同一个目录(同卷)


# --------------------------------------------------------------------------- #
# 工具函数
# --------------------------------------------------------------------------- #
def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _safe_time(date_time) -> Tuple[int, int, int, int, int, int]:
    try:
        year, month, day, hour, minute, second = date_time
        if year < 1980:
            return (1980, 1, 1, 0, 0, 0)
        return (year, month, day, hour, minute, second)
    except Exception:
        return (1980, 1, 1, 0, 0, 0)


def _next_rid(used) -> str:
    index = 1
    while f"rId{index}" in used:
        index += 1
    return f"rId{index}"


def _base_dir(rels_name: str) -> str:
    """关系部件对应的源部件所在目录。

    ``_rels/.rels``                    → ``""``
    ``word/_rels/document.xml.rels``   → ``"word"``
    ``customXml/_rels/item1.xml.rels`` → ``"customXml"``
    """
    return posixpath.dirname(posixpath.dirname(rels_name.replace("\\", "/")))


def _resolve_target(rels_name: str, target: str) -> str:
    target = target.replace("\\", "/").strip()
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(_base_dir(rels_name), target))


def _is_external(rel) -> bool:
    return (rel.get("TargetMode") or "").lower() == "external"


def _looks_external(target: str) -> bool:
    low = target.lower()
    return ("://" in low or low.startswith("file:")
            or low.startswith("mailto:") or low.startswith("#"))


def _quiet_remove(path: str) -> None:
    if not path:
        return
    try:
        if os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


# --------------------------------------------------------------------------- #
# OPC 包结构修复
# --------------------------------------------------------------------------- #
def _fix_relationships(name: str, data: bytes, present: set,
                       lower_present: set, strip_macros: bool,
                       report: PackageReport) -> bytes:
    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        raise RepairError(f"{name} 无法解析: {exc}") from exc

    used = set()
    id_sigs: dict = {}
    changed = False
    for rel in list(root):
        if _local(rel.tag) != "Relationship":
            continue

        rel_type = (rel.get("Type") or "").lower()
        if strip_macros and ("/vbaproject" in rel_type or "/vbadata" in rel_type):
            root.remove(rel)
            report.removed_macro_rels.append(f"{name}: {rel.get('Target')}")
            changed = True
            continue

        rid = (rel.get("Id") or "").strip()
        sig = (rel.get("Type") or "", rel.get("Target") or "",
               rel.get("TargetMode") or "")
        if rid and rid in id_sigs:
            if sig in id_sigs[rid]:
                # 同一 Id、同一目标:病毒复制出来的冗余关系,直接删掉
                root.remove(rel)
                report.duplicate_rels.append(f"{name}: {rid}")
                changed = True
                continue
            # 同一 Id 指向不同目标:保留,但换一个未占用的 Id
            new_id = _next_rid(used)
            rel.set("Id", new_id)
            report.renamed_rel_ids.append(f"{name}: {rid} → {new_id}")
            rid = new_id
            changed = True
        elif not rid:
            new_id = _next_rid(used)
            rel.set("Id", new_id)
            report.renamed_rel_ids.append(f"{name}: (空) → {new_id}")
            rid = new_id
            changed = True
        used.add(rid)
        id_sigs.setdefault(rid, set()).add(sig)

        target = (rel.get("Target") or "").strip()
        if not target:
            root.remove(rel)
            report.dangling_rels.append(f"{name}: (空目标)")
            changed = True
            continue
        if _is_external(rel) or _looks_external(target):
            continue

        resolved = _resolve_target(name, target)
        if resolved not in present and unquote(resolved) not in present \
                and resolved.lower() not in lower_present:
            root.remove(rel)
            report.dangling_rels.append(f"{name}: {target}")
            changed = True

    if not changed:
        return data
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _fix_content_types(data: bytes, present: set, lower_present: set,
                       report: PackageReport) -> bytes:
    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        raise RepairError(f"[Content_Types].xml 无法解析: {exc}") from exc

    changed = False
    seen_parts = set()
    seen_ext = set()
    for node in list(root):
        tag = _local(node.tag)
        if tag == "Override":
            part = (node.get("PartName") or "").lstrip("/")
            key = part.lower()
            if key in seen_parts or key not in lower_present:
                root.remove(node)
                report.removed_overrides.append(f"[Content_Types]: {part}")
                changed = True
                continue
            seen_parts.add(key)
        elif tag == "Default":
            ext = (node.get("Extension") or "").lower()
            if ext in seen_ext:
                root.remove(node)
                changed = True
                continue
            seen_ext.add(ext)

    if not changed:
        return data
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def sanitize_package(src: str, dst: Optional[str] = None, strip_macros: bool = True,
                     log=None) -> PackageReport:
    """把损坏的 OPC 包重写成结构合法的包(不动正文内容)。

    处理项:
      * 同名重复部件 → 只保留第一份;
      * 关系文件内重复/缺失的 Id → 重新分配未占用的 rIdN;
      * 指向不存在部件的内部关系 → 删除;
      * ``strip_macros=True`` → 移除 word/vbaProject.bin 等宏部件及其关系;
      * [Content_Types].xml 里重复或指向缺失部件的 Override → 删除。

    ``dst`` 省略时原地替换。返回修复明细。
    """
    src = os.path.abspath(src)
    dst = os.path.abspath(dst) if dst else src
    report = PackageReport()

    infos = []
    names = set()
    try:
        zf = zipfile.ZipFile(src, "r")
    except (OSError, zipfile.BadZipFile) as exc:
        raise RepairError(f"不是合法的 OPC/zip 包: {exc}") from exc
    with zf:
        for info in zf.infolist():
            name = info.filename
            if name.endswith("/"):
                continue
            if name in names:
                report.duplicate_parts.append(name)
                continue
            names.add(name)
            infos.append(info)

    # 丢弃宏部件(以及它们的 rel 由 _fix_relationships 处理)
    if strip_macros:
        kept = []
        for info in infos:
            if info.filename.lower() in MACRO_PARTS:
                report.removed_parts.append(info.filename)
                names.discard(info.filename)
                continue
            kept.append(info)
        infos = kept

    present = set(names)
    lower_present = {n.lower() for n in names}

    def _rewrite(name: str, data: bytes) -> bytes:
        low = name.lower()
        if low.endswith(".rels"):
            return _fix_relationships(name, data, present, lower_present,
                                      strip_macros, report)
        if low == "[content_types].xml":
            return _fix_content_types(data, present, lower_present, report)
        return data

    # 写临时文件再原子替换,避免写一半损坏目标。除 .rels / Content_Types
    # 之外的部件按块流式拷贝,几十 MB/上百 MB 的图片不会整体读进内存。
    directory = os.path.dirname(dst) or "."
    fd, tmp = tempfile.mkstemp(prefix=".d2d-sanitize-", suffix=".tmp", dir=directory)
    os.close(fd)
    try:
        with zipfile.ZipFile(src, "r") as zin, \
                zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for info in infos:
                name = info.filename
                zinfo = zipfile.ZipInfo(name, date_time=_safe_time(info.date_time))
                zinfo.compress_type = zipfile.ZIP_DEFLATED
                zinfo.external_attr = info.external_attr
                low = name.lower()
                if low.endswith(".rels") or low == "[content_types].xml":
                    out.writestr(zinfo, _rewrite(name, zin.read(info)))
                else:
                    with zin.open(info) as fsrc, out.open(zinfo, "w") as fdst:
                        shutil.copyfileobj(fsrc, fdst, 1024 * 1024)
        os.replace(tmp, dst)
    except Exception:
        _quiet_remove(tmp)
        raise

    if log and report.changed:
        log(f"  包结构修复: {report.summary()}")
    return report


# --------------------------------------------------------------------------- #
# Word 重存
# --------------------------------------------------------------------------- #
def _new_temp_path(directory: str, suffix: str = ".docm") -> str:
    fd, path = tempfile.mkstemp(prefix=TEMP_PREFIX, suffix=suffix, dir=directory)
    os.close(fd)
    os.remove(path)          # 让 Word 自己新建,避免“文件已存在”提示
    return path


def _work_directory(src: str) -> Tuple[str, bool]:
    """优先在源文件同目录放临时文件(可原子覆盖),不行再退回系统临时目录。"""
    src_dir = os.path.dirname(os.path.abspath(src))
    try:
        probe = _new_temp_path(src_dir, ".probe")
        os.remove(probe)
        return src_dir, True
    except OSError:
        return tempfile.gettempdir(), False


def _validate_docm(path: str) -> Tuple[bool, str]:
    if not os.path.isfile(path) or os.path.getsize(path) == 0:
        return False, "修复产物不存在或为空"
    ok, problems = validate_package(path)
    if not ok:
        return False, "；".join(problems[:3])
    if not has_part(path, "word/document.xml"):
        return False, "修复产物缺少 word/document.xml"
    return True, ""


def _word_save_docm(word_app, open_path: str, save_path: str,
                    open_and_repair: bool) -> Tuple[bool, str]:
    """用 Word 打开文档并另存为 .docm。返回 (是否成功, 说明)。"""
    _quiet_remove(save_path)
    doc = None
    try:
        doc = word_app.Documents.Open(
            FileName=os.path.abspath(open_path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=open_and_repair,
            NoEncodingDialog=True,
        )
        saver = getattr(doc, "SaveAs2", None) or doc.SaveAs
        saver(FileName=os.path.abspath(save_path),
              FileFormat=WD_FORMAT_XML_DOCUMENT_MACRO,
              AddToRecentFiles=False)
        return True, ""
    except Exception as exc:
        return False, str(exc).replace("\r", " ").replace("\n", " ")
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass


def repair_document(word_app, src: str, strip_macros: bool = True,
                    log=None) -> RepairResult:
    """尝试修复一个无法直接打开的 .docm。

    返回的 ``RepairResult.path`` 是修复后的临时 .docm;调用方负责把它
    覆盖回原文件(或转换后清理)。
    """
    src = os.path.abspath(src)
    if not os.path.isfile(src):
        return RepairResult(False, detail="源文件不存在")
    if os.path.getsize(src) == 0:
        return RepairResult(False, detail="源文件为空")

    work_dir, in_place = _work_directory(src)
    candidate = _new_temp_path(work_dir)
    success = False
    notes: List[str] = []

    try:
        # ---- 第一步:Word 自带的“打开并修复” ------------------------- #
        if log:
            log("  修复: 尝试 Word 自带的“打开并修复”…")
        ok, detail = _word_save_docm(word_app, src, candidate, open_and_repair=True)
        if ok:
            ok, why = _validate_docm(candidate)
            if ok:
                if strip_macros and has_macros(candidate):
                    if log:
                        log("  修复: 去掉修复产物中的 VBA 宏工程…")
                    sanitize_package(candidate, candidate, strip_macros=True)
                    ok, why = _validate_docm(candidate)
                if ok:
                    success = True
                    return RepairResult(True, candidate, "word",
                                        "Word 自带修复成功", notes, in_place)
            if log:
                log(f"  修复: Word 修复产物未通过校验({why})")
        else:
            if log:
                log(f"  修复: Word 自带修复未成功({detail})")

        _quiet_remove(candidate)

        # ---- 第二步:修 OPC 包结构,再让 Word 重存 -------------------- #
        normalized = _new_temp_path(work_dir)
        try:
            try:
                report = sanitize_package(src, normalized, strip_macros=strip_macros)
            except RepairError as exc:
                return RepairResult(False, detail=f"包结构修复失败: {exc}")
            notes.append(report.summary())
            if log:
                log(f"  修复: Word 修复无效,改修 OPC 包结构({report.summary()})…")
            ok, detail = _word_save_docm(word_app, normalized, candidate,
                                         open_and_repair=False)
            if not ok:
                return RepairResult(False, detail=f"包结构修复后仍无法用 Word 打开: {detail}")
            ok, why = _validate_docm(candidate)
            if not ok:
                return RepairResult(False, detail=f"Word 重存后校验失败: {why}")
            if strip_macros and has_macros(candidate):
                sanitize_package(candidate, candidate, strip_macros=True)
                ok, why = _validate_docm(candidate)
                if not ok:
                    return RepairResult(False, detail=f"去宏后校验失败: {why}")
            success = True
            return RepairResult(True, candidate, "package",
                                "已修正包结构并用 Word 重新另存为 .docm", notes, in_place)
        finally:
            _quiet_remove(normalized)
    except Exception as exc:                     # 兜底:修复失败不能让整批中断
        return RepairResult(False, detail=f"修复过程异常: {exc}")
    finally:
        if not success:
            _quiet_remove(candidate)


# --------------------------------------------------------------------------- #
# 覆盖原文件
# --------------------------------------------------------------------------- #
def replace_source(src: str, repaired: str) -> Tuple[bool, str]:
    """用修复后的 .docm 覆盖原文件。返回 (是否成功, 说明)。"""
    src = os.path.abspath(src)
    repaired = os.path.abspath(repaired)
    if src == repaired:
        return True, ""
    if not os.path.isfile(repaired):
        return False, "修复产物不存在"
    try:
        os.chmod(src, stat.S_IWRITE)          # 清掉只读属性
    except OSError:
        pass
    try:
        os.replace(repaired, src)             # 同卷:原子替换
        return True, ""
    except OSError as exc:
        try:
            shutil.copyfile(repaired, src)    # 跨卷/占用等原因:退回复制
            _quiet_remove(repaired)
            return True, ""
        except OSError as exc2:
            return False, f"{exc}；复制回原位置也失败: {exc2}"
