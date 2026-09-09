"""修订模式（Track Changes）插入：以 ``<w:ins>`` 标记向 Word 文档插入段落 / 表格。

用途：把新数据以「待接受」的修订形式插入到已有 ``.docx`` 中，可在 Word 审阅视图里
逐条接受 / 拒绝，而不是直接改写正文。核心实现是在目标元素外包一层 ``<w:ins>``。

约定：

- ``<w:ins>`` 的 ``w:id`` 全文档唯一，必须避开原文档已有的修订 id（``next_revision_id``）。
- 插入内容默认不覆盖、不删除原文档任何内容，只做增量插入。
- 若 ``track_changes=False``，则退化为普通插入（不产生修订标记）。
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def next_revision_id(doc: Document) -> int:
    """扫描文档中已有的全部 ``w:id``，返回可用的起始 id（max+1）。

    Word 要求修订标记的 id 全文档唯一；重复会导致打开时提示修复。
    """
    max_id = 0
    for el in doc.element.iter():
        v = el.get(qn("w:id"))
        if v is not None:
            try:
                max_id = max(max_id, int(v))
            except ValueError:
                pass
    return max_id + 1


def _norm(s: str) -> str:
    """折叠所有空白，便于稳定匹配标题文本。"""
    return re.sub(r"\s+", "", s or "")


def find_heading(
    doc: Document, text: str, style_prefix: Optional[str] = None
) -> Optional[Paragraph]:
    """按标题文本查找段落。

    优先匹配「样式名同时包含 ``style_prefix`` 与『标题』」的段落（用于自定义样式，
    如 ``"GF报告"``）；找不到时兜底按纯文本前缀匹配任意段落。
    """
    target = _norm(text)
    if style_prefix:
        for p in doc.paragraphs:
            style = p.style.name if p.style else ""
            if style.startswith(style_prefix) and "标题" in style:
                if _norm(p.text).startswith(target):
                    return p
    for p in doc.paragraphs:
        if _norm(p.text).startswith(target):
            return p
    return None


class TrackedInsert:
    """修订插入器：统一作者 / 时间，自动分配不冲突的 ``w:id``。

    用法：先构造一个实例，再反复调用 ``paragraph_before`` / ``table_before``
    在指定锚点之前插入内容；每条插入都会包进 ``<w:ins>`` 修订标记。
    """

    def __init__(
        self,
        doc: Document,
        author: str = "修订",
        date: str = "2026-09-09T10:00:00Z",
        enabled: bool = True,
    ):
        self.doc = doc
        self.author = author
        self.date = date
        self.enabled = enabled
        self._next = next_revision_id(doc)

    def paragraph_before(
        self,
        anchor: Paragraph,
        text: str,
        *,
        bold: bool = False,
        style: Optional[str] = None,
    ) -> Paragraph:
        """在 anchor 之前插入一个段落，并包进 ``<w:ins>`` 修订标记。"""
        p = anchor.insert_paragraph_before()
        if style:
            try:
                p.style = self.doc.styles[style]
            except Exception:
                pass
        run = p.add_run(text)
        run.bold = bold
        self._wrap(p._p)
        return p

    def table_before(self, anchor: Paragraph, rows: List[List[str]]) -> Any:
        """在 anchor 之前插入一个表格（list of list of str），并包进 ``<w:ins>``。"""
        if not rows:
            return None
        ncols = max(len(r) for r in rows)
        tbl = self.doc.add_table(rows=0, cols=ncols)
        try:
            tbl.style = "Table Grid"
        except Exception:
            pass
        for row in rows:
            cells = tbl.add_row().cells
            for i, val in enumerate(row):
                cells[i].text = str(val)
        anchor._p.addprevious(tbl._tbl)
        self._wrap(tbl._tbl)
        return tbl

    def _wrap(self, element) -> None:
        """把已定位到目标位置的元素包进 ``<w:ins>``（lxml 的 append 会重挂载）。"""
        if not self.enabled:
            return
        ins = OxmlElement("w:ins")
        ins.set(qn("w:id"), str(self._next))
        self._next += 1
        ins.set(qn("w:author"), self.author)
        ins.set(qn("w:date"), self.date)
        element.addprevious(ins)
        ins.append(element)


def merge_into_docx(
    src: str,
    dst: str,
    insertions: List[Dict[str, Any]],
    *,
    author: str = "修订",
    date: str = "2026-09-09T10:00:00Z",
    body_style: Optional[str] = None,
    style_prefix: Optional[str] = None,
    track_changes: bool = True,
) -> Dict[str, Any]:
    """把多条插入项按锚点合并进文档并另存，返回统计结果。

    参数：

    - ``insertions``：插入项列表，每项字段：
      ``anchor``（必填，标题文本，用于定位插入点）、``label``（可选，加粗标签段）、
      ``body``（可选，正文段）、``table``（可选，list of list of str）。
    - ``style_prefix``：标题样式前缀（如 ``"GF报告"``），用于更精确地定位标题；
      省略时按纯文本匹配。
    - ``track_changes``：True 时插入内容以修订标记呈现，可在 Word 中接受 / 拒绝。

    返回 ``{"total": n, "done": n, "skipped": [锚点...]}``。
    """
    doc = Document(src)
    ti = TrackedInsert(doc, author=author, date=date, enabled=track_changes)

    result: Dict[str, Any] = {"total": len(insertions), "done": 0, "skipped": []}
    for item in insertions:
        anchor = find_heading(doc, item.get("anchor", ""), style_prefix=style_prefix)
        if anchor is None:
            result["skipped"].append(item.get("anchor", ""))
            continue
        if item.get("label"):
            ti.paragraph_before(anchor, item["label"], bold=True, style=body_style)
        if item.get("body"):
            ti.paragraph_before(anchor, item["body"], style=body_style)
        if item.get("table"):
            ti.table_before(anchor, item["table"])
        result["done"] += 1

    doc.save(dst)
    return result
