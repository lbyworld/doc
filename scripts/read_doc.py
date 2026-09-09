#!/usr/bin/env python
"""便捷脚本：读取 .docx。

用法：
    python scripts/read_doc.py input.docx [-f markdown|text|json] [-o out.md]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docwork.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main(["read", *sys.argv[1:]]))
