#!/usr/bin/env python
"""便捷脚本：从 Markdown / JSON 撰写 .docx。

用法：
    python scripts/write_doc.py input.md -o output.docx
    python scripts/write_doc.py input.json -o output.docx [--title "标题"]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docwork.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main(["write", *sys.argv[1:]]))
