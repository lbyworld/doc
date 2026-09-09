#!/usr/bin/env python
"""便捷脚本：按 JSON 计划以修订模式合并内容进 .docx。

用法：
    python scripts/merge_docx.py 计划.json --src 源.docx --dst 输出.docx

计划 JSON 的结构（author/date/body_style/heading_style_prefix/track_changes/insertions）
见示例 docs/plans/third_party_merge.json。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docwork.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main(["merge", *sys.argv[1:]]))
