# -*- coding: utf-8 -*-
"""OOXML(OPC zip)解析:提取正文文本、结构统计、包完整性检查。

.docm 与 .docx 都是 OOXML 包,内部结构几乎一致,唯一差别是 .docm 多一个
word/vbaProject.bin(宏工程)。因此"文本 + 结构"逐项对比可以有效判断
转换过程是否丢内容。
"""
from __future__ import annotations

import os
import zipfile
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from .constants import NS_W

W = "{%s}" % NS_W


class OoxmlError(Exception):
    """OOXML 包无法读取 / 解析。"""


def _open_zip(path: str) -> zipfile.ZipFile:
    if not os.path.isfile(path):
        raise OoxmlError(f"文件不存在: {path}")
    if os.path.getsize(path) == 0:
        raise OoxmlError("文件大小为 0(空文件)")
    try:
        return zipfile.ZipFile(path, "r")
    except zipfile.BadZipFile as exc:
        raise OoxmlError(f"不是合法的 zip/OOXML 包: {exc}") from exc


def list_parts(path: str) -> List[str]:
    """列出包内全部部件名。"""
    with _open_zip(path) as zf:
        return zf.namelist()


def read_part(path: str, part: str) -> bytes:
    with _open_zip(path) as zf:
        try:
            return zf.read(part)
        except KeyError as exc:
            raise OoxmlError(f"缺少部件: {part}") from exc


def has_part(path: str, part: str) -> bool:
    try:
        with _open_zip(path) as zf:
            return part in zf.namelist()
    except OoxmlError:
        return False


def _parse_document_xml(path: str) -> ET.Element:
    raw = read_part(path, "word/document.xml")
    try:
        return ET.fromstring(raw)
    except ET.ParseError as exc:
        raise OoxmlError(f"word/document.xml 解析失败: {exc}") from exc


def _text_of(element: ET.Element) -> str:
    """提取元素内的可见文本(w:t),并把制表/换行还原成字符。"""
    chunks: List[str] = []
    for node in element.iter():
        tag = node.tag
        if tag == W + "t":
            chunks.append(node.text or "")
        elif tag == W + "tab":
            chunks.append("\t")
        elif tag in (W + "br", W + "cr"):
            chunks.append("\n")
        elif tag == W + "noBreakHyphen":
            chunks.append("-")
    return "".join(chunks)


def extract_body_text(path: str) -> str:
    """提取正文(document.xml)的纯文本。"""
    return _text_of(_parse_document_xml(path))


def extract_text_of_parts(path: str, prefix: str) -> str:
    """提取所有以 prefix 开头的部件(如 word/header)中的文本。"""
    chunks: List[str] = []
    with _open_zip(path) as zf:
        for name in sorted(zf.namelist()):
            if name.startswith(prefix) and name.endswith(".xml"):
                try:
                    root = ET.fromstring(zf.read(name))
                except ET.ParseError:
                    continue
                chunks.append(_text_of(root))
    return "".join(chunks)


def paragraph_texts(path: str) -> List[str]:
    """按文档顺序返回每一段的文本,用于逐段比对(比整篇比对更能定位差异)。"""
    root = _parse_document_xml(path)
    return [_text_of(para) for para in root.iter(W + "p")]


def structure_stats(path: str) -> Dict[str, int]:
    """统计文档结构元素数量,用于 docm/docx 逐项比对。

    注意:runs 属于"序列化实现细节"。Word 另存时会重排 run(例如把
    <w:fldSimple> 简写域展开成 fldChar/instrText 复杂域),run 数变化
    并不代表内容变化,因此校验时只作提示,不作为失败依据。
    """
    root = _parse_document_xml(path)
    stats = {
        "paragraphs": 0,
        "runs": 0,
        "tables": 0,
        "rows": 0,
        "cells": 0,
        "drawings": 0,
        "pictures": 0,
        "hyperlinks": 0,
        "fields": 0,
        "simple_fields": 0,
        "complex_fields": 0,
        "bookmarks": 0,
        "sections": 0,
        "footnote_refs": 0,
        "endnote_refs": 0,
        "chars": 0,
    }
    for node in root.iter():
        tag = node.tag
        if tag == W + "p":
            stats["paragraphs"] += 1
        elif tag == W + "r":
            stats["runs"] += 1
        elif tag == W + "tbl":
            stats["tables"] += 1
        elif tag == W + "tr":
            stats["rows"] += 1
        elif tag == W + "tc":
            stats["cells"] += 1
        elif tag == W + "drawing":
            stats["drawings"] += 1
        elif tag == W + "pict":
            stats["pictures"] += 1
        elif tag == W + "hyperlink":
            stats["hyperlinks"] += 1
        elif tag == W + "fldSimple":
            stats["fields"] += 1
            stats["simple_fields"] += 1
        elif tag == W + "instrText":
            stats["fields"] += 1
            stats["complex_fields"] += 1
        elif tag == W + "bookmarkStart":
            stats["bookmarks"] += 1
        elif tag == W + "sectPr":
            stats["sections"] += 1
        elif tag == W + "footnoteReference":
            stats["footnote_refs"] += 1
        elif tag == W + "endnoteReference":
            stats["endnote_refs"] += 1
    stats["chars"] = len(extract_body_text(path))
    return stats


def media_files(path: str) -> List[str]:
    """列出 word/media 下的媒体文件(图片等)。"""
    with _open_zip(path) as zf:
        return sorted(n for n in zf.namelist() if n.startswith("word/media/"))


def part_counts(path: str) -> Dict[str, int]:
    """统计各类部件数量(页眉/页脚/批注/脚注 …)。"""
    with _open_zip(path) as zf:
        names = zf.namelist()
    return {
        "headers": len([n for n in names if n.startswith("word/header") and n.endswith(".xml")]),
        "footers": len([n for n in names if n.startswith("word/footer") and n.endswith(".xml")]),
        "footnotes": int("word/footnotes.xml" in names),
        "endnotes": int("word/endnotes.xml" in names),
        "comments": int("word/comments.xml" in names),
        "embeddings": len([n for n in names if n.startswith("word/embeddings/")]),
        "media": len([n for n in names if n.startswith("word/media/")]),
    }


def has_macros(path: str) -> bool:
    """包内是否还残留 VBA 工程。"""
    try:
        with _open_zip(path) as zf:
            names = zf.namelist()
    except OoxmlError:
        return False
    return any("vbaProject" in n for n in names)


def validate_package(path: str) -> Tuple[bool, List[str]]:
    """检查包结构是否合法。返回 (是否通过, 问题列表)。"""
    problems: List[str] = []
    try:
        with _open_zip(path) as zf:
            bad = zf.testzip()
            if bad is not None:
                problems.append(f"CRC 校验失败的部件: {bad}")
            names = zf.namelist()
            for required in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
                if required not in names:
                    problems.append(f"缺少必需部件: {required}")
            for name in names:
                if name.endswith(".xml") or name.endswith(".rels"):
                    try:
                        ET.fromstring(zf.read(name))
                    except ET.ParseError as exc:
                        problems.append(f"XML 解析失败 {name}: {exc}")
    except OoxmlError as exc:
        return False, [str(exc)]
    return (not problems), problems
