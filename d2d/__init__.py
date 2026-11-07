# -*- coding: utf-8 -*-
"""d2d — DOCM 批量无损转换为 DOCX 的工具包。

子模块:
    constants  Word/OOXML 常量
    scanner    磁盘 / 文件夹 / 单文件扫描
    ooxml      OOXML(OPC zip)结构解析与文本提取
    engine     Word COM 转换引擎
    repair     损坏 .docm 的修复(Word 修复 + OPC 包结构修复)
    testgen    测试用 .docm 生成
    verify     转换结果机器校验
    gui        Tkinter 图形界面
    cli        命令行 / 自检入口
"""

__version__ = "1.0.0"
__all__ = [
    "constants",
    "scanner",
    "ooxml",
    "engine",
    "repair",
    "testgen",
    "verify",
    "gui",
    "cli",
]
