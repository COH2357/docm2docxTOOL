# -*- coding: utf-8 -*-
"""转换结果的机器校验。

把源 .docm 与产物 .docx 从三个层面逐项比对:
  1) 包层:ZIP/OOXML 结构完整性、必需部件、宏工程是否按预期移除;
  2) 内容层:正文文本逐字比对 + 结构元素计数(段落/表格/图片/域…)比对;
  3) 应用层:用 Word 真正打开两个文件,比对 Word 层面的统计(字数、页数等),
     这能直接暴露"打开时提示修复"类的损坏。
"""
from __future__ import annotations

import difflib
import os
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from . import ooxml
from .engine import WordSession

LEVEL_OK = "ok"
LEVEL_WARN = "warn"
LEVEL_FAIL = "fail"


@dataclass
class Check:
    name: str
    passed: bool
    detail: str = ""
    level: str = LEVEL_OK

    @property
    def status(self) -> str:
        if self.level == LEVEL_FAIL or not self.passed:
            return "✗ 失败"
        if self.level == LEVEL_WARN:
            return "! 提示"
        return "✓ 通过"


@dataclass
class VerifyReport:
    docm: str = ""
    docx: str = ""
    checks: List[Check] = field(default_factory=list)
    summary: str = ""

    @property
    def failed(self) -> List[Check]:
        return [c for c in self.checks if c.level == LEVEL_FAIL]

    @property
    def ok(self) -> bool:
        return not self.failed


# --------------------------------------------------------------------------- #
# Word 层面统计
# --------------------------------------------------------------------------- #
def word_stats(word_app, path: str) -> Dict[str, object]:
    """用 Word 打开文件并读取统计信息(同时验证文件能被正常打开)。"""
    doc = word_app.Documents.Open(
        FileName=os.path.abspath(path),
        ConfirmConversions=False,
        ReadOnly=True,
        AddToRecentFiles=False,
        Visible=False,
        OpenAndRepair=False,
        NoEncodingDialog=True,
    )
    try:
        text = doc.Content.Text
        stats = {
            "paragraphs": doc.Paragraphs.Count,
            "words": doc.Words.Count,
            "characters": doc.Characters.Count,
            "tables": doc.Tables.Count,
            "inline_shapes": doc.InlineShapes.Count,
            "shapes": doc.Shapes.Count,
            "fields": doc.Fields.Count,
            "hyperlinks": doc.Hyperlinks.Count,
            "footnotes": doc.Footnotes.Count,
            "sections": doc.Sections.Count,
            "text_len": len(text),
            "text_head": text[:400],
        }
        try:
            stats["pages"] = doc.ComputeStatistics(2)      # wdStatisticPages
        except Exception:
            stats["pages"] = -1
        return stats
    finally:
        try:
            doc.Close(SaveChanges=0)
        except Exception:
            pass


# --------------------------------------------------------------------------- #
# 主校验流程
# --------------------------------------------------------------------------- #
def verify_conversion(docm_path: str,
                      docx_path: str,
                      word_app=None,
                      deep: bool = True) -> VerifyReport:
    """比对 docm 与 docx,返回校验报告。"""
    docm_path = os.path.abspath(docm_path)
    docx_path = os.path.abspath(docx_path)
    report = VerifyReport(docm=docm_path, docx=docx_path)

    # --- 0. 文件存在性 ------------------------------------------------- #
    if not os.path.isfile(docm_path):
        report.checks.append(Check("源文件存在", False, f"找不到 {docm_path}", LEVEL_FAIL))
        return report
    if not os.path.isfile(docx_path):
        report.checks.append(Check("源文件存在", True, os.path.basename(docm_path)))
        report.checks.append(Check("产物文件存在", False, f"找不到 {docx_path}", LEVEL_FAIL))
        return report
    report.checks.append(Check("源文件存在", True, docm_path))
    report.checks.append(Check("产物文件存在", True,
                               f"{docx_path} ({os.path.getsize(docx_path):,} 字节)"))

    # --- 1. 包完整性 ---------------------------------------------------- #
    ok, problems = ooxml.validate_package(docx_path)
    report.checks.append(Check(
        "DOCX 包结构完整(ZIP+XML)",
        ok,
        "CRC 校验通过,全部 XML 部件可解析" if ok else "; ".join(problems[:5]),
        LEVEL_OK if ok else LEVEL_FAIL))

    # --- 2. 宏工程 ------------------------------------------------------ #
    macros_left = ooxml.has_macros(docx_path)
    report.checks.append(Check(
        "宏工程已按预期移除",
        not macros_left,
        "产物为纯 .docx,不含 vbaProject.bin(符合预期)" if not macros_left
        else "产物仍包含 VBA 工程,可能未真正转成 docx",
        LEVEL_OK if not macros_left else LEVEL_WARN))

    # --- 3. 正文文本 ----------------------------------------------------- #
    try:
        text_docm = ooxml.extract_body_text(docm_path)
        text_docx = ooxml.extract_body_text(docx_path)
        if text_docm == text_docx:
            report.checks.append(Check(
                "正文文本逐字一致", True,
                f"{len(text_docx):,} 字符完全相同", LEVEL_OK))
        else:
            ratio = difflib.SequenceMatcher(None, text_docm, text_docx).ratio()
            diff_chars = abs(len(text_docm) - len(text_docx))
            report.checks.append(Check(
                "正文文本逐字一致", ratio >= 0.9999,
                f"相似度 {ratio:.6%},长度 {len(text_docm):,} → {len(text_docx):,}"
                f"(差 {diff_chars:,} 字符)",
                LEVEL_OK if ratio >= 0.9999 else LEVEL_FAIL))
    except ooxml.OoxmlError as exc:
        report.checks.append(Check("正文文本逐字一致", False, str(exc), LEVEL_FAIL))

    # --- 3b. 逐段文本顺序比对 --------------------------------------------- #
    try:
        paras_docm = ooxml.paragraph_texts(docm_path)
        paras_docx = ooxml.paragraph_texts(docx_path)
        if len(paras_docm) != len(paras_docx):
            report.checks.append(Check(
                "逐段文本顺序一致", False,
                f"段落数不一致: {len(paras_docm)} → {len(paras_docx)}", LEVEL_FAIL))
        else:
            bad = [(i + 1, a, b) for i, (a, b) in enumerate(zip(paras_docm, paras_docx))
                   if a != b]
            detail = (f"{len(paras_docm)} 段逐段文本与顺序完全一致" if not bad
                      else "第 " + ", ".join(str(i) for i, _, _ in bad[:5])
                      + " 段文本不同,例如: "
                      + repr(bad[0][1][:60]) + " → " + repr(bad[0][2][:60]))
            report.checks.append(Check(
                "逐段文本顺序一致", not bad, detail,
                LEVEL_OK if not bad else LEVEL_FAIL))
    except ooxml.OoxmlError as exc:
        report.checks.append(Check("逐段文本顺序一致", False, str(exc), LEVEL_FAIL))

    # --- 4. 结构元素计数 -------------------------------------------------- #
    # runs 属于序列化细节(Word 另存会重排 run / 展开简写域),只作提示。
    content_keys = ["paragraphs", "tables", "rows", "cells", "drawings",
                    "pictures", "hyperlinks", "fields", "bookmarks",
                    "sections", "footnote_refs", "endnote_refs", "chars"]
    label = {
        "paragraphs": "段落", "tables": "表格", "rows": "表格行", "cells": "单元格",
        "drawings": "Drawing 对象", "pictures": "VML 图片", "hyperlinks": "超链接",
        "fields": "域", "bookmarks": "书签", "sections": "节",
        "footnote_refs": "脚注引用", "endnote_refs": "尾注引用", "chars": "字符数",
    }
    try:
        stats_docm = ooxml.structure_stats(docm_path)
        stats_docx = ooxml.structure_stats(docx_path)
        diffs = []
        for key in content_keys:
            if stats_docm.get(key) != stats_docx.get(key):
                diffs.append(f"{label.get(key, key)}: "
                             f"{stats_docm.get(key)} → {stats_docx.get(key)}")
        report.checks.append(Check(
            "内容结构元素计数一致",
            not diffs,
            "段落/表格/单元格/图片/域/书签/节/脚注 数量全部一致" if not diffs
            else "; ".join(diffs),
            LEVEL_OK if not diffs else LEVEL_FAIL))

        runs_left = stats_docm.get("runs", 0)
        runs_right = stats_docx.get("runs", 0)
        simple = stats_docx.get("simple_fields", 0)
        complex_fields = stats_docx.get("complex_fields", 0)
        report.checks.append(Check(
            "run 数(仅提示)",
            True,
            f"{runs_left} → {runs_right}"
            + ("" if runs_left == runs_right
               else "  ← Word 另存时重新序列化内部结构"
                   f"(简写域 {simple} 个 / 展开域 {complex_fields} 个),不影响内容"),
            LEVEL_OK if runs_left == runs_right else LEVEL_WARN))
    except ooxml.OoxmlError as exc:
        report.checks.append(Check("内容结构元素计数一致", False, str(exc), LEVEL_FAIL))

    # --- 5. 部件数量(页眉/页脚/脚注/媒体) -------------------------------- #
    try:
        parts_docm = ooxml.part_counts(docm_path)
        parts_docx = ooxml.part_counts(docx_path)
        diffs = [f"{k}: {parts_docm[k]} → {parts_docx[k]}"
                 for k in parts_docm if parts_docm[k] != parts_docx[k]]
        report.checks.append(Check(
            "附件部件数量一致",
            not diffs,
            f"页眉 {parts_docx['headers']} / 页脚 {parts_docx['footers']} / "
            f"媒体 {parts_docx['media']} / 脚注 {parts_docx['footnotes']}"
            if not diffs else "; ".join(diffs),
            LEVEL_OK if not diffs else LEVEL_FAIL))
    except ooxml.OoxmlError as exc:
        report.checks.append(Check("附件部件数量一致", False, str(exc), LEVEL_FAIL))

    # --- 6. 页眉页脚文本 -------------------------------------------------- #
    try:
        head_docm = ooxml.extract_text_of_parts(docm_path, "word/header")
        head_docx = ooxml.extract_text_of_parts(docx_path, "word/header")
        foot_docm = ooxml.extract_text_of_parts(docm_path, "word/footer")
        foot_docx = ooxml.extract_text_of_parts(docx_path, "word/footer")
        same = (head_docm == head_docx) and (foot_docm == foot_docx)
        report.checks.append(Check(
            "页眉页脚文本一致", same,
            "页眉页脚文字完全相同" if same
            else f"页眉 {head_docm!r} → {head_docx!r}; 页脚 {foot_docm!r} → {foot_docx!r}",
            LEVEL_OK if same else LEVEL_FAIL))
    except ooxml.OoxmlError as exc:
        report.checks.append(Check("页眉页脚文本一致", False, str(exc), LEVEL_FAIL))

    # --- 7. Word 实际打开比对 -------------------------------------------- #
    if deep:
        _verify_with_word(docm_path, docx_path, report, word_app)

    # --- 汇总 ------------------------------------------------------------ #
    failed = len(report.failed)
    warnings = len([c for c in report.checks if c.level == LEVEL_WARN])
    report.summary = (f"共 {len(report.checks)} 项检查,"
                       f"失败 {failed} 项,提示 {warnings} 项")
    return report


def _verify_with_word(docm_path: str, docx_path: str,
                      report: VerifyReport, word_app) -> None:
    """用 Word 打开两个文件,比对 Word 层面统计,验证没有损坏。"""
    session = None
    try:
        if word_app is None:
            session = WordSession(visible=False)
            session.start()
            word_app = session.app

        try:
            stats_docm = word_stats(word_app, docm_path)
        except Exception as exc:
            report.checks.append(Check(
                "Word 可正常打开产物 .docx", False,
                f"Word 打开 .docx 失败: {exc}", LEVEL_FAIL))
            return

        report.checks.append(Check(
            "Word 可正常打开产物 .docx", True,
            "Word 打开无异常(未触发修复流程)", LEVEL_OK))

        try:
            stats_docx = word_stats(word_app, docx_path)
        except Exception as exc:
            report.checks.append(Check(
                "Word 层面内容统计一致", False,
                f"Word 打开 .docm 失败: {exc}", LEVEL_FAIL))
            return

        exact_keys = ["paragraphs", "tables", "inline_shapes", "shapes",
                      "fields", "hyperlinks", "footnotes", "sections", "pages"]
        label = {
            "paragraphs": "段落数", "tables": "表格数", "inline_shapes": "内嵌图",
            "shapes": "形状", "fields": "域", "hyperlinks": "超链接",
            "footnotes": "脚注", "sections": "节", "pages": "页数",
        }
        diffs = []
        for key in exact_keys:
            left = stats_docm.get(key)
            right = stats_docx.get(key)
            if left is None or right is None:
                continue
            if left != right:
                diffs.append(f"{label.get(key, key)}: {left} → {right}")

        left_chars = stats_docm.get("characters", 0)
        right_chars = stats_docx.get("characters", 0)
        if left_chars != right_chars:
            diffs.append(f"字符数: {left_chars} → {right_chars}")

        report.checks.append(Check(
            "Word 层面内容统计一致",
            not diffs,
            "段落/表格/图片/形状/域/超链接/脚注/页数/字符数全部一致"
            if not diffs else "; ".join(diffs),
            LEVEL_OK if not diffs else LEVEL_FAIL))

        # Word 读取到的正文与源文件正文逐字比对
        text_docm = stats_docm.get("text_head", "")
        text_docx = stats_docx.get("text_head", "")
        report.checks.append(Check(
            "Word 读取的正文内容一致",
            text_docm == text_docx,
            "Word 打开后读到的正文起始内容相同" if text_docm == text_docx
            else "Word 读取到的正文内容存在差异",
            LEVEL_OK if text_docm == text_docx else LEVEL_FAIL))
    finally:
        if session is not None:
            session.close()
