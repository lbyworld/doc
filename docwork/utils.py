"""公共工具函数：标题样式识别、内联 Markdown 解析、中文字体设置。"""

from __future__ import annotations

import re
from typing import List, Tuple

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt


# ---------- 标题样式 ----------

def is_heading(style_name: str) -> bool:
    """判断段落样式是否为标题（兼容中英文样式名）。"""
    return style_name.startswith("Heading") or style_name.startswith("标题")


def heading_level(style_name: str) -> int:
    """从样式名提取标题级别，默认返回 1。"""
    match = re.search(r"(\d+)", style_name)
    return int(match.group(1)) if match else 1


# ---------- 内联 Markdown 解析 ----------

# 每个片段为 (text, bold, italic, code)
Inline = Tuple[str, bool, bool, bool]

_INLINE_RE = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)")


def parse_inline(text: str) -> List[Inline]:
    """把内联 Markdown 文本解析为 (text, bold, italic, code) 片段列表。

    支持 ``**加粗**``、``*斜体*``、``` `行内代码` ```。
    """
    parts: List[Inline] = []
    for seg in _INLINE_RE.split(text):
        if not seg:
            continue
        if seg.startswith("**") and seg.endswith("**") and len(seg) >= 4:
            parts.append((seg[2:-2], True, False, False))
        elif seg.startswith("`") and seg.endswith("`") and len(seg) >= 2:
            parts.append((seg[1:-1], False, False, True))
        elif seg.startswith("*") and seg.endswith("*") and len(seg) >= 2:
            parts.append((seg[1:-1], False, True, False))
        else:
            parts.append((seg, False, False, False))
    return parts


def add_inline_runs(paragraph, text: str) -> None:
    """把带内联 Markdown 的文本按格式写入段落。"""
    for seg, bold, italic, code in parse_inline(text):
        run = paragraph.add_run(seg)
        run.bold = bold
        run.italic = italic
        if code:
            run.font.name = "Consolas"


# ---------- 文档默认样式 ----------

def set_normal_font(doc: Document, latin: str = "Calibri", east_asia: str = "宋体") -> None:
    """设置 Normal 样式的西文与中文字体、字号（11pt）。"""
    style = doc.styles["Normal"]
    style.font.name = latin
    style.font.size = Pt(11)
    try:
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn("w:eastAsia"), east_asia)
    except Exception:
        # 某些模板可能没有 rPr，忽略中文字体设置
        pass
