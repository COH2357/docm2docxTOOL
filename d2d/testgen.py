# -*- coding: utf-8 -*-
"""生成用于验证转换无损性的测试 .docm 文档。

文档里刻意塞进各种容易在格式转换中丢失的元素:
多级标题 / 段落 / 项目符号 / 编号 / 表格 / 图片 / 形状 / 超链接 /
域 / 页眉页脚 / 脚注 / 分页符 / 字符级格式 / 中英文混排 / 特殊字符。
"""
from __future__ import annotations

import os
import struct
import zlib
from datetime import datetime
from typing import Dict, List

from .constants import (
    WD_ALIGN_PARAGRAPH_CENTER, WD_ALIGN_PARAGRAPH_JUSTIFY,
    WD_ALIGN_PARAGRAPH_LEFT, WD_FIELD_DATE, WD_FIELD_NUM_PAGES,
    WD_FIELD_PAGE, WD_INLINE_SHAPE_PICTURE, WD_PAGE_BREAK,
)

# 内置样式常量(用数值而非名称,避免中英文 Word 差异)
WD_STYLE_NORMAL = -1
WD_STYLE_HEADING1 = -2
WD_STYLE_HEADING2 = -3
WD_STYLE_HEADING3 = -4
WD_STYLE_TITLE = -63
WD_STYLE_CAPTION = -35

WD_HEADER_FOOTER_PRIMARY = 1
WD_COLLAPSE_START = 1
MSO_SHAPE_RECTANGLE = 1
WD_COLOR_RED = 0x0000FF          # WdColor 为 BGR 顺序
WD_COLOR_DARK_BLUE = 0x800000
WD_HIGHLIGHT_YELLOW = 7
WD_UNDERLINE_SINGLE = 1


def rgb(r: int, g: int, b: int) -> int:
    """Word 的 WdColor 是 BGR 排列。"""
    return r + (g << 8) + (b << 16)


# --------------------------------------------------------------------------- #
# 生成测试用 PNG(不依赖 Pillow)
# --------------------------------------------------------------------------- #
def make_test_png(path: str, width: int = 160, height: int = 100) -> str:
    """写一张渐变棋盘图,供 Word 插入用。"""
    rows = []
    for y in range(height):
        row = bytearray([0])                      # filter type 0
        for x in range(width):
            r = (x * 255) // max(1, width - 1)
            g = (y * 255) // max(1, height - 1)
            b = 255 - r
            if (x // 10 + y // 10) % 2 == 0:
                r, g, b = r // 2, g // 2, b // 2
            row += bytes((r & 0xFF, g & 0xFF, b & 0xFF))
        rows.append(bytes(row))
    raw = b"".join(rows)

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    blob = (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(raw, 9))
            + chunk(b"IEND", b""))
    with open(path, "wb") as handle:
        handle.write(blob)
    return path


# --------------------------------------------------------------------------- #
# 文档构建辅助
# --------------------------------------------------------------------------- #
def _append_paragraph(doc, text: str = "", style: int = None):
    """在文档末尾追加一段并返回该段落对象。"""
    paragraphs = doc.Paragraphs
    last = paragraphs(paragraphs.Count)
    if last.Range.Text.strip("\r\x07\x0c") == "":
        para = last
    else:
        doc.Content.InsertParagraphAfter()
        para = doc.Paragraphs(doc.Paragraphs.Count)
    if text:
        para.Range.InsertBefore(text)
    if style is not None:
        try:
            doc.Paragraphs(doc.Paragraphs.Count).Range.Style = doc.Styles(style)
        except Exception:
            pass
    return doc.Paragraphs(doc.Paragraphs.Count)


def _tail_range(doc):
    """文档末尾的折叠 Range(用于插图/表格/域)。"""
    end = doc.Content.End - 1
    if end < 0:
        end = 0
    return doc.Range(end, end)


def _para_range(doc, index: int):
    """第 index 段(1-based)的 Range。"""
    return doc.Paragraphs(index).Range


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def create_test_docm(word_app, docm_path: str) -> Dict[str, object]:
    """生成测试 .docm,返回已写入元素清单。"""
    docm_path = os.path.abspath(docm_path)
    os.makedirs(os.path.dirname(docm_path), exist_ok=True)

    png_path = os.path.join(os.path.dirname(docm_path), "_test_image.png")
    make_test_png(png_path)

    created: List[str] = []
    notes: List[str] = []
    doc = word_app.Documents.Add()

    def step(name: str, func) -> None:
        try:
            func()
            created.append(name)
        except Exception as exc:
            notes.append(f"{name} 写入失败: {exc}")

    try:
        # 1. 标题 + 元信息 ---------------------------------------------- #
        step("标题(Title 样式)", lambda: _append_paragraph(
            doc, "DOCM → DOCX 无损转换测试文档", WD_STYLE_TITLE))
        step("元信息段落", lambda: _append_paragraph(
            doc, f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  |  "
                 f"生成器: DOCM→DOCX 转换器自检模块", WD_STYLE_NORMAL))
        step("空行分隔", lambda: _append_paragraph(doc, "", WD_STYLE_NORMAL))

        # 2. 多级标题 ------------------------------------------------------ #
        step("一级标题", lambda: _append_paragraph(
            doc, "一、多级标题与样式", WD_STYLE_HEADING1))
        step("二级标题", lambda: _append_paragraph(
            doc, "1.1 二级标题示例", WD_STYLE_HEADING2))
        step("三级标题", lambda: _append_paragraph(
            doc, "1.1.1 三级标题示例", WD_STYLE_HEADING3))
        step("正文段落(两端对齐)", lambda: _append_paragraph(
            doc, "这是一段用于校验的正文内容,包含中文、English mixed text、"
                 "数字 1234567890、标点符号(，。；：？！)以及特殊字符 "
                 "& < > \" ' © ® ™ § ¶ → ← ± ½ ¼ ¾ № ℃ € £ ¥ ¢。"
                 "转换前后这段文字应逐字一致。", WD_STYLE_NORMAL))

        # 3. 列表 ---------------------------------------------------------- #
        step("项目符号列表标题", lambda: _append_paragraph(
            doc, "二、列表", WD_STYLE_HEADING1))
        bullets = ["项目符号第一条:Alpha", "项目符号第二条:Beta",
                   "项目符号第三条:Gamma(含数字 123)"]
        for text in bullets:
            step(f"项目符号 {text[:6]}", lambda t=text: _bullet(doc, t))
        numbers = ["编号第一条:One", "编号第二条:Two", "编号第三条:Three"]
        for text in numbers:
            step(f"编号 {text[:6]}", lambda t=text: _numbered(doc, t))

        # 4. 表格 ---------------------------------------------------------- #
        step("表格标题", lambda: _append_paragraph(
            doc, "三、表格", WD_STYLE_HEADING1))
        step("表格 3x4", lambda: _add_table(doc))
        step("表格 2x3", lambda: _add_small_table(doc))
        step("表格说明", lambda: _append_paragraph(
            doc, "表格说明:上表为 3 行 4 列,下表为 2 行 3 列,含中文与数字混合内容。",
            WD_STYLE_NORMAL))

        # 5. 图片 ---------------------------------------------------------- #
        step("图片标题", lambda: _append_paragraph(
            doc, "四、图片与图形", WD_STYLE_HEADING1))
        step("插入 PNG 图片", lambda: _add_picture(doc, png_path))
        step("图片题注", lambda: _append_paragraph(
            doc, "图 1 测试用渐变棋盘图(PNG)", WD_STYLE_CAPTION))
        step("插入矩形形状", lambda: _add_shape(doc))

        # 6. 超链接 / 域 --------------------------------------------------- #
        step("超链接标题", lambda: _append_paragraph(
            doc, "五、超链接与域", WD_STYLE_HEADING1))
        step("插入超链接", lambda: _add_hyperlink(doc))
        step("插入日期域", lambda: _add_field(doc, WD_FIELD_DATE))
        step("插入页码域", lambda: _add_field(doc, WD_FIELD_PAGE))
        step("插入总页数域", lambda: _add_field(doc, WD_FIELD_NUM_PAGES))

        # 7. 页眉页脚 ------------------------------------------------------ #
        step("页眉", lambda: _set_header_footer(doc))

        # 8. 分页符 -------------------------------------------------------- #
        step("分页符", lambda: _add_page_break(doc))
        step("第二页标题", lambda: _append_paragraph(
            doc, "六、第二页内容(分页符之后)", WD_STYLE_HEADING1))

        # 9. 字符格式 ------------------------------------------------------ #
        step("字符格式段落", lambda: _format_run(doc))

        # 10. 脚注 --------------------------------------------------------- #
        step("脚注", lambda: _add_footnote(doc))

        # 11. 结尾 --------------------------------------------------------- #
        step("结尾段落", lambda: _append_paragraph(
            doc, "— 测试文档结束,以上内容应完整保留在 .docx 中 —",
            WD_STYLE_NORMAL))

        # 保存为 docm ------------------------------------------------------ #
        saver = getattr(doc, "SaveAs2", None) or doc.SaveAs
        saver(FileName=docm_path, FileFormat=13, AddToRecentFiles=False)
    finally:
        try:
            doc.Close(SaveChanges=0)
        except Exception:
            pass

    return {
        "docm": docm_path,
        "created": created,
        "notes": notes,
        "png": png_path,
    }


# --------------------------------------------------------------------------- #
# 具体元素写入
# --------------------------------------------------------------------------- #
def _bullet(doc, text: str) -> None:
    para = _append_paragraph(doc, text, WD_STYLE_NORMAL)
    try:
        para.Range.ListFormat.ApplyBulletDefault()
    except Exception:
        pass


def _numbered(doc, text: str) -> None:
    para = _append_paragraph(doc, text, WD_STYLE_NORMAL)
    try:
        para.Range.ListFormat.ApplyNumberDefault()
    except Exception:
        pass


def _add_table(doc) -> None:
    para = _append_paragraph(doc, "", WD_STYLE_NORMAL)
    rng = para.Range
    rng.Collapse(WD_COLLAPSE_START)
    table = doc.Tables.Add(rng, 3, 4)
    data = [
        ["表头 A1", "表头 A2", "表头 A3", "表头 A4"],
        ["数据 B1", "数值 42", "文本 & <ok>", "日期 2026-01-01"],
        ["数据 C1", "数值 43", "文本 100%", "日期 2026-06-30"],
    ]
    for r, row in enumerate(data, 1):
        for c, value in enumerate(row, 1):
            try:
                table.Cell(r, c).Range.Text = value
            except Exception:
                pass
    try:
        table.Rows(1).Range.Font.Bold = True
        table.Rows(1).HeadingFormat = True
    except Exception:
        pass


def _add_small_table(doc) -> None:
    para = _append_paragraph(doc, "", WD_STYLE_NORMAL)
    rng = para.Range
    rng.Collapse(WD_COLLAPSE_START)
    table = doc.Tables.Add(rng, 2, 3)
    for r, row in enumerate([["小表 1", "小表 2", "小表 3"], ["值 1", "值 2", "值 3"]], 1):
        for c, value in enumerate(row, 1):
            try:
                table.Cell(r, c).Range.Text = value
            except Exception:
                pass


def _add_picture(doc, png_path: str) -> None:
    para = _append_paragraph(doc, "", WD_STYLE_NORMAL)
    rng = para.Range
    rng.Collapse(WD_COLLAPSE_START)
    doc.InlineShapes.AddPicture(FileName=png_path, LinkToFile=False,
                                SaveWithDocument=True, Range=rng)


def _add_shape(doc) -> None:
    para = _append_paragraph(doc, "", WD_STYLE_NORMAL)
    rng = para.Range
    rng.Collapse(WD_COLLAPSE_START)
    shape = doc.Shapes.AddShape(Type=MSO_SHAPE_RECTANGLE, Left=60, Top=60,
                                Width=220, Height=70, Anchor=rng)
    try:
        shape.TextFrame.TextRange.Text = "形状文本框 Shape Text"
        shape.Fill.ForeColor.RGB = rgb(230, 240, 255)
        shape.TextFrame.TextRange.Font.Size = 12
    except Exception:
        pass


def _add_hyperlink(doc) -> None:
    para = _append_paragraph(doc, "", WD_STYLE_NORMAL)
    rng = para.Range
    rng.Collapse(WD_COLLAPSE_START)
    doc.Hyperlinks.Add(Anchor=rng, Address="https://www.example.com/docm2docx",
                       TextToDisplay="示例超链接 Example Link")


def _add_field(doc, field_type: int) -> None:
    para = _append_paragraph(doc, "", WD_STYLE_NORMAL)
    rng = para.Range
    rng.Collapse(WD_COLLAPSE_START)
    doc.Fields.Add(Range=rng, Type=field_type, PreserveFormatting=True)


def _set_header_footer(doc) -> None:
    section = doc.Sections(1)
    try:
        section.Headers(WD_HEADER_FOOTER_PRIMARY).Range.Text = \
            "页眉: DOCM→DOCX 无损转换测试文档"
    except Exception:
        pass
    try:
        section.Footers(WD_HEADER_FOOTER_PRIMARY).Range.Text = \
            "页脚: 第 PAGE 页 / 共 NUMPAGES 页"
    except Exception:
        pass


def _add_page_break(doc) -> None:
    para = _append_paragraph(doc, "", WD_STYLE_NORMAL)
    rng = para.Range
    rng.Collapse(WD_COLLAPSE_START)
    rng.InsertBreak(WD_PAGE_BREAK)


def _format_run(doc) -> None:
    para = _append_paragraph(
        doc, "字符格式示例:粗体 斜体 下划线 彩色 加大字号 高亮", WD_STYLE_NORMAL)
    try:
        rng = para.Range
        rng.Font.Bold = True
        rng.Font.Size = 14
        rng.Font.Color = WD_COLOR_RED
        rng.Font.HighlightColorIndex = WD_HIGHLIGHT_YELLOW
        rng.Font.Underline = WD_UNDERLINE_SINGLE
    except Exception:
        pass
    para2 = _append_paragraph(
        doc, "居中段落示例(Centered)", WD_STYLE_NORMAL)
    try:
        para2.Alignment = WD_ALIGN_PARAGRAPH_CENTER
    except Exception:
        pass


def _add_footnote(doc) -> None:
    para = _append_paragraph(doc, "本段落带有一个脚注引用。", WD_STYLE_NORMAL)
    try:
        rng = para.Range
        rng.MoveEnd(1, -1)
        doc.Footnotes.Add(Range=rng, Text="脚注内容:脚注也应完整保留在 docx 中。")
    except Exception:
        pass
