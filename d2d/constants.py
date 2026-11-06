# -*- coding: utf-8 -*-
"""Word COM / OOXML 相关常量。"""

# --- Word 文档格式 (WdSaveFormat) ---
WD_FORMAT_DOCUMENT_97 = 0
WD_FORMAT_TEMPLATE_97 = 1
WD_FORMAT_TEXT = 2
WD_FORMAT_RTF = 6
WD_FORMAT_HTML = 8
WD_FORMAT_XML_DOCUMENT = 12          # .docx (Word 2007 XML)
WD_FORMAT_XML_DOCUMENT_MACRO = 13    # .docm (启用宏的 XML 文档)
WD_FORMAT_XML_TEMPLATE = 14
WD_FORMAT_XML_TEMPLATE_MACRO = 15
WD_FORMAT_DOCUMENT_DEFAULT = 16      # .docx (Word 默认格式)
WD_FORMAT_PDF = 17

# --- 关闭文档时的保存选项 (WdSaveOptions) ---
WD_DO_NOT_SAVE_CHANGES = 0
WD_SAVE_CHANGES = -1
WD_PROMPT_TO_SAVE_CHANGES = -2

# --- 警告级别 (WdAlertLevel) ---
WD_ALERTS_NONE = 0
WD_ALERTS_MESSAGE_BOX = 1
WD_ALERTS_ALL = 2

# --- Office 自动化安全 (MsoAutomationSecurity) ---
MSO_AUTOMATION_SECURITY_LOW = 1
MSO_AUTOMATION_SECURITY_BY_UI = 2
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3   # 禁用宏,不弹任何提示(最安全)

# --- 字段类型 (WdFieldType),测试文档用 ---
WD_FIELD_DATE = 31
WD_FIELD_TIME = 32
WD_FIELD_PAGE = 33
WD_FIELD_NUM_PAGES = 26
WD_FIELD_FILE_NAME = 29

# --- 分节符 / 换行符 ---
WD_PAGE_BREAK = 7
WD_SECTION_BREAK_NEXT_PAGE = 2

# --- 对齐方式 ---
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2
WD_ALIGN_PARAGRAPH_JUSTIFY = 3

# --- 列表模板 ---
WD_LIST_BULLET = 1
WD_LIST_NUMBER = 2

# --- 内联图形 ---
WD_INLINE_SHAPE_PICTURE = 3

# --- 文件后缀 ---
DOCM_SUFFIX = ".docm"
DOCX_SUFFIX = ".docx"
TEMP_FILE_PREFIX = "~$"

# --- OOXML 命名空间 ---
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# --- 覆盖策略 ---
OVERWRITE_SKIP = "skip"
OVERWRITE_REPLACE = "overwrite"
OVERWRITE_RENAME = "rename"

OVERWRITE_LABELS = {
    OVERWRITE_SKIP: "跳过(保留已有文件)",
    OVERWRITE_REPLACE: "覆盖(删除已有文件)",
    OVERWRITE_RENAME: "自动重命名(name (1).docx)",
}
OVERWRITE_BY_LABEL = {label: key for key, label in OVERWRITE_LABELS.items()}

# --- 源文件处理策略(转换成功后) ---
CLEANUP_NONE = "none"          # 不动源文件
CLEANUP_RECYCLE = "recycle"    # 移到回收站(可恢复)
CLEANUP_DELETE = "delete"      # 永久删除(不可恢复)

CLEANUP_LABELS = {
    CLEANUP_NONE: "不处理(保留原 .docm)",
    CLEANUP_RECYCLE: "移到回收站(可恢复,推荐)",
    CLEANUP_DELETE: "永久删除(不可恢复!)",
}
CLEANUP_BY_LABEL = {label: key for key, label in CLEANUP_LABELS.items()}
