"""读取 Word 文档：按文档顺序提取段落 / 标题 / 表格，输出 markdown / text / json。"""

from __future__ import annotations

import json
from typing import Any, Dict, List

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

from .utils import heading_level, is_heading


def _iter_block_items(doc):
    """按文档出现顺序遍历段落与表格（python-docx 官方推荐方式）。"""
    from docx.oxml.ns import qn

    parent = doc.element.body
    for child in parent.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def read_docx(path: str) -> List[Dict[str, Any]]:
    """读取 .docx，返回结构化块列表。

    每个块为一个字典，``type`` 取值为：

    - ``heading``：``{"type": "heading", "level": int, "text": str}``
    - ``paragraph``：``{"type": "paragraph", "text": str}``
    - ``table``：``{"type": "table", "rows": [[str, ...], ...]}``

    该结构可直接交给 ``structured_to_docx`` 反向生成文档。
    """
    doc = Document(path)
    blocks: List[Dict[str, Any]] = []
    for item in _iter_block_items(doc):
        if isinstance(item, Paragraph):
            style = item.style.name if item.style else ""
            if is_heading(style):
                blocks.append(
                    {"type": "heading", "level": heading_level(style), "text": item.text}
                )
            elif item.text.strip():
                # 列表样式（List Bullet / List Number）按普通段落提取，保留文本
                blocks.append({"type": "paragraph", "text": item.text})
        elif isinstance(item, Table):
            blocks.append(
                {"type": "table", "rows": [[c.text for c in r.cells] for r in item.rows]}
            )
    return blocks


def _rows_to_markdown(rows: List[List[str]]) -> str:
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    padded = [list(r) + [""] * (width - len(r)) for r in rows]
    header = padded[0]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    for row in padded[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def docx_to_markdown(path: str) -> str:
    """读取 .docx，输出 Markdown 文本。"""
    lines: List[str] = []
    for b in read_docx(path):
        t = b["type"]
        if t == "heading":
            lines.append("#" * b["level"] + " " + b["text"])
        elif t == "paragraph":
            lines.append(b["text"])
        elif t == "table":
            lines.append(_rows_to_markdown(b["rows"]))
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def docx_to_text(path: str) -> str:
    """读取 .docx，输出纯文本（表格以 | 分隔单元格）。"""
    lines: List[str] = []
    for b in read_docx(path):
        if b["type"] == "table":
            for row in b["rows"]:
                lines.append(" | ".join(row))
        else:
            lines.append(b["text"])
    return "\n".join(lines)


def docx_to_json(path: str) -> str:
    """读取 .docx，输出 JSON 文本（结构化块列表）。"""
    return json.dumps(read_docx(path), ensure_ascii=False, indent=2)
