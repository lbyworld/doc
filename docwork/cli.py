"""命令行入口：``python -m docwork read|write|edit``。"""

from __future__ import annotations

import argparse
import json
import sys

from . import editor, reader, revisions, writer


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="docwork",
        description="Word 文档 (.docx) 读取 / 修改 / 撰写工具",
    )
    sub = p.add_subparsers(dest="command", required=True)

    # read
    r = sub.add_parser("read", help="读取 .docx，输出 markdown / text / json")
    r.add_argument("input", help="输入的 .docx 文件")
    r.add_argument("-f", "--format", choices=["markdown", "text", "json"], default="markdown")
    r.add_argument("-o", "--output", help="输出文件路径（默认打印到标准输出）")

    # write
    w = sub.add_parser("write", help="从 Markdown 或 JSON 撰写 .docx")
    w.add_argument("input", help="输入的 .md 或 .json 文件")
    w.add_argument("-o", "--output", required=True, help="输出的 .docx 文件")
    w.add_argument("--title", help="文档标题（仅 JSON 输入时可覆盖）")

    # edit
    e = sub.add_parser("edit", help="修改 .docx")
    e.add_argument("input", help="输入的 .docx 文件")
    e.add_argument("-o", "--output", required=True, help="输出的 .docx 文件")
    e.add_argument("--replace", action="append", metavar="OLD=NEW",
                   help="查找替换（可多次指定）")
    e.add_argument("--append", dest="append_text", help="在文末追加段落")
    e.add_argument("--prepend", dest="prepend_text", help="在文首插入段落")
    e.add_argument("--remove", dest="remove_text", help="删除所有包含该文本的段落")

    # merge
    m = sub.add_parser("merge", help="按 JSON 计划以修订模式合并内容进 .docx")
    m.add_argument("plan", help="合并计划 JSON 文件（含 author/date/style/insertions）")
    m.add_argument("--src", required=True, help="源 .docx 文件")
    m.add_argument("--dst", required=True, help="输出的 .docx 文件")
    return p


def _cmd_read(args) -> int:
    if args.format == "markdown":
        out = reader.docx_to_markdown(args.input)
    elif args.format == "text":
        out = reader.docx_to_text(args.input)
    else:
        out = reader.docx_to_json(args.input)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(out)
        print(f"已输出到 {args.output}", file=sys.stderr)
    else:
        print(out)
    return 0


def _cmd_write(args) -> int:
    with open(args.input, encoding="utf-8") as f:
        text = f.read()
    if args.input.lower().endswith(".json"):
        data = json.loads(text)
        writer.structured_to_docx(data, args.output, title=args.title)
    else:
        writer.markdown_to_docx(text, args.output)
    print(f"已生成 {args.output}", file=sys.stderr)
    return 0


def _cmd_edit(args) -> int:
    replacements = {}
    if args.replace:
        for r in args.replace:
            if "=" not in r:
                print(f"错误：--replace 需为 OLD=NEW 格式，收到 {r!r}", file=sys.stderr)
                return 2
            old, new = r.split("=", 1)
            replacements[old] = new

    editor.edit_docx(
        args.input,
        args.output,
        replacements=replacements or None,
        append=args.append_text,
        prepend=args.prepend_text,
        remove=args.remove_text,
    )
    print(f"已保存 {args.output}", file=sys.stderr)
    return 0


def _cmd_merge(args) -> int:
    with open(args.plan, encoding="utf-8") as f:
        plan = json.load(f)
    result = revisions.merge_into_docx(
        args.src,
        args.dst,
        plan.get("insertions", []),
        author=plan.get("author", "修订"),
        date=plan.get("date", "2026-09-09T10:00:00Z"),
        body_style=plan.get("body_style"),
        style_prefix=plan.get("heading_style_prefix"),
        track_changes=plan.get("track_changes", True),
    )
    print(
        f"已保存 {args.dst}：完成 {result['done']}/{result['total']} 处插入。",
        file=sys.stderr,
    )
    if result["skipped"]:
        print(f"未找到锚点（已跳过）：{result['skipped']}", file=sys.stderr)
    return 0


def main(argv=None) -> int:
    # 强制 stdout/stderr 使用 UTF-8，避免 Windows 下重定向输出变成 GBK 乱码。
    # 注意：直接在 GBK 终端（cmd）中查看仍需先执行 `chcp 65001`。
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass

    args = _build_parser().parse_args(argv)
    if args.command == "read":
        return _cmd_read(args)
    if args.command == "write":
        return _cmd_write(args)
    if args.command == "edit":
        return _cmd_edit(args)
    if args.command == "merge":
        return _cmd_merge(args)
    return 1
