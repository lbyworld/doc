"""docwork — Word 文档 (.docx) 的读取 / 修改 / 撰写工具库。

三个核心能力：
- 读取：``read_docx`` / ``docx_to_markdown`` / ``docx_to_text`` / ``docx_to_json``
- 撰写：``markdown_to_docx`` / ``structured_to_docx``
- 修改：``replace_text`` / ``append_paragraph`` / ``prepend_paragraph`` /
        ``remove_paragraphs_matching`` / ``edit_docx``

也可作为命令行工具使用：``python -m docwork read|write|edit ...``
"""

__version__ = "0.1.0"

from .reader import (
    read_docx,
    docx_to_markdown,
    docx_to_text,
    docx_to_json,
)
from .writer import markdown_to_docx, structured_to_docx
from .editor import (
    replace_text,
    append_paragraph,
    prepend_paragraph,
    remove_paragraphs_matching,
    edit_docx,
)
from .revisions import merge_into_docx, TrackedInsert, find_heading, next_revision_id

__all__ = [
    "read_docx",
    "docx_to_markdown",
    "docx_to_text",
    "docx_to_json",
    "markdown_to_docx",
    "structured_to_docx",
    "replace_text",
    "append_paragraph",
    "prepend_paragraph",
    "remove_paragraphs_matching",
    "edit_docx",
    "merge_into_docx",
    "TrackedInsert",
    "find_heading",
    "next_revision_id",
]
