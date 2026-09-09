#!/usr/bin/env python
"""便捷脚本：修改 .docx。

用法：
    python scripts/edit_doc.py input.docx -o output.docx --replace 旧=新 [--replace ...]
    python scripts/edit_doc.py input.docx -o output.docx --append "文末追加"
    python scripts/edit_doc.py input.docx -o output.docx --prepend "文首插入"
    python scripts/edit_doc.py input.docx -o output.docx --remove "含此文本的段落"
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docwork.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main(["edit", *sys.argv[1:]]))
