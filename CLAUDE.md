# docwork — Word 文档读写改工作流

> 在 `d:\claude\doc` 目录下，对 Word 文档（`.docx`）进行**读取、修改、撰写**的 Python 工具集。

## 项目概述

本仓库提供一套针对 `.docx` 文件的三步工作流，核心是一个可导入的 Python 包 `docwork`，同时提供命令行入口与便捷脚本：

1. **读取（Read）** — 把 `.docx` 提取为 Markdown / 纯文本 / 结构化 JSON
2. **撰写（Write）** — 从 Markdown 或结构化 JSON 生成 `.docx`
3. **修改（Edit）** — 对 `.docx` 做查找替换、增删段落
4. **修订模式合并（Merge）** — 按 JSON 计划以 Word 修订模式（Track Changes）插入新内容

依赖仅有 `python-docx`（+ 其依赖 `lxml`），无 Pandoc 等外部工具。

## Python 环境（重要）

**本项目的 Python 位于 Anaconda 的 `docx` 环境，不要使用系统默认 Python。**

- 环境路径：`D:\anaconda3\envs\docx`
- 解释器：`D:\anaconda3\envs\docx\python.exe`（Python 3.12.14）
- 已装依赖：`python-docx 1.2.0`、`lxml 6.1.2`

运行方式（二选一）：

```bash
# 方式 A：直接调用完整路径（最稳妥，推荐在脚本/命令中使用）
D:/anaconda3/envs/docx/python.exe -m docwork ...

# 方式 B：激活环境后调用
conda activate docx
python -m docwork ...
```

> 提示：`D:\anaconda3\envs\` 下另有 `yololabel` 环境，与本项目无关，勿误用。

## 项目结构

```
d:\claude\doc\
├── CLAUDE.md                  # 本文档（给 Claude Code 的工作指南）
├── README.md                  # 使用说明
├── requirements.txt           # 依赖（python-docx / lxml）
├── .gitignore
├── docwork/                   # 核心 Python 包
│   ├── __init__.py            # 公开 API（导入入口）
│   ├── __main__.py            # 支持 python -m docwork
│   ├── cli.py                 # 命令行入口（read/write/edit/merge 子命令）
│   ├── reader.py              # 读取：docx → markdown/text/json
│   ├── writer.py              # 撰写：markdown/json → docx
│   ├── editor.py              # 修改：查找替换、增删段落
│   ├── revisions.py           # 修订模式：<w:ins> 插入段落/表格、按计划合并
│   └── utils.py               # 公共工具：样式识别、内联 Markdown 解析、字体
├── scripts/                   # 便捷脚本（免 -m，直接跑）
│   ├── read_doc.py
│   ├── write_doc.py
│   ├── edit_doc.py
│   └── merge_docx.py          # 按 JSON 计划以修订模式合并内容
├── docs/
│   ├── input/                 # 待处理的 .docx 输入
│   ├── output/                # 处理结果输出（已 gitignore）
│   ├── plans/                 # 合并计划 JSON（数据与逻辑分离，可复用）
│   └── templates/             # 文档模板
└── examples/
    └── sample.md              # 示例 Markdown（含标题/列表/表格/代码块）
```

## 命令速查表

所有命令在**项目根目录** `d:\claude\doc` 下执行，Python 用 `D:/anaconda3/envs/docx/python.exe`（下称 `$PY`）。

### 读取

```bash
$PY -m docwork read 输入.docx                     # 输出 Markdown 到终端
$PY -m docwork read 输入.docx -o 输出.md          # 输出 Markdown 到文件
$PY -m docwork read 输入.docx -f text             # 纯文本
$PY -m docwork read 输入.docx -f json -o 输出.json # 结构化 JSON
```

### 撰写

```bash
$PY -m docwork write 输入.md  -o 输出.docx         # Markdown → docx
$PY -m docwork write 输入.json -o 输出.docx        # JSON → docx（可用 --title）
```

### 修改

```bash
$PY -m docwork edit 输入.docx -o 输出.docx --replace 旧=新          # 查找替换
$PY -m docwork edit 输入.docx -o 输出.docx --replace a=1 --replace b=2  # 多次替换
$PY -m docwork edit 输入.docx -o 输出.docx --append "文末追加"       # 追加段落
$PY -m docwork edit 输入.docx -o 输出.docx --prepend "文首插入"      # 前插段落
$PY -m docwork edit 输入.docx -o 输出.docx --remove "含此文本"       # 删除匹配段落
```

### 修订模式合并（Track Changes）

按 JSON 计划把新增内容以 `<w:ins>` 修订标记插入到指定章节锚点之前（可在 Word 审阅视图中逐条接受/拒绝）：

```bash
$PY -m docwork merge 计划.json --src 源.docx --dst 输出.docx
# 本次试验的完整计划（5 处第三方检测数据）：
$PY -m docwork merge docs/plans/third_party_merge.json \
    --src "docs/input/10-合同编号-宇勘复芯-技术总结报告0901(1).docx" \
    --dst "docs/output/技术总结报告_补充第三方测试数据_修订模式.docx"
```

等价便捷脚本（免 `-m`）：`$PY scripts/read_doc.py` / `write_doc.py` / `edit_doc.py` / `merge_docx.py`，参数相同。

## 代码调用示例

```python
import docwork

# 读取为结构化块列表
blocks = docwork.read_docx("a.docx")
print(docwork.docx_to_markdown("a.docx"))

# 撰写
docwork.markdown_to_docx("# 标题\n\n正文", "b.docx")
docwork.structured_to_docx(blocks, "roundtrip.docx")   # 块列表可直接反向生成

# 修改
docwork.edit_docx("a.docx", "c.docx",
                  replacements={"张三": "李四"},
                  append="追加内容", remove="待删除")

# 修订模式合并（按锚点插入，内容以修订标记呈现）
docwork.merge_into_docx(
    "源.docx", "输出.docx",
    [{"anchor": "2.6 解决的关键技术及技术途径",
      "label": "（5）第三方检测机构算力测试结果",
      "body": "三次均值 6.243492 TOPS，满足 ≥6 TOPS 要求。",
      "table": [["次数", "算力/TOPS"], ["1", "6.158032"]]}],
    author="第三方检测", body_style="GF报告正文",
    style_prefix="GF报告", track_changes=True)
```

## 脚本与模块清单

| 脚本 / 模块 | 作用 |
| --- | --- |
| `docwork/reader.py` | 读取：`.docx` → 结构化块 / Markdown / 纯文本 / JSON |
| `docwork/writer.py` | 撰写：Markdown / JSON → `.docx` |
| `docwork/editor.py` | 修改：查找替换、追加 / 前插 / 删除段落 |
| `docwork/revisions.py` | 修订模式插入：`<w:ins>` 标记段落 / 表格，按 JSON 计划合并 |
| `docwork/cli.py` | 命令行入口：`read` / `write` / `edit` / `merge` 子命令 |
| `docwork/utils.py` | 公共工具：标题样式识别、内联 Markdown 解析、中文字体 |
| `scripts/read_doc.py` | 读取便捷脚本（等价 `python -m docwork read`） |
| `scripts/write_doc.py` | 撰写便捷脚本（等价 `python -m docwork write`） |
| `scripts/edit_doc.py` | 修改便捷脚本（等价 `python -m docwork edit`） |
| `scripts/merge_docx.py` | 修订模式合并便捷脚本（等价 `python -m docwork merge`） |
| `docs/plans/*.json` | 合并计划（数据与逻辑分离，新增合并只需写 JSON 计划） |

## 结构化块格式（读写的中间表示）

`read_docx` 返回的块列表与 `structured_to_docx` 的输入保持一致，可无损往返：

```json
[
  {"type": "heading",   "level": 1, "text": "标题"},
  {"type": "paragraph", "text": "正文"},
  {"type": "table",     "rows": [["a","b"],["1","2"]]},
  {"type": "list",      "ordered": false, "items": ["项1","项2"]}
]
```

## 关键实现约定

- **跨 run 替换**：python-docx 会把一段文本拆成多个 `run`，`editor._replace_in_paragraph` 先合并全文再替换，替换后**该段落 run 的原有格式会被重置为第一个 run 的格式**（这是简化取舍）。
- **标题识别**：`reader` 通过段落样式名（`Heading 1` / `标题 1`）判断标题级别，兼容中英文样式名。
- **内联格式**：`writer` 支持 `**加粗**`、`*斜体*`、`` `行内代码` ``；但 `reader` 目前**不会**把 run 的粗体/斜体还原成 Markdown 标记（读取为纯文本），这是已知限制。
- **中文字体**：`utils.set_normal_font` 设置 Normal 样式西文 Calibri + 中文宋体，避免中文默认字体异常。
- **修订模式插入**：`revisions.TrackedInsert` 在插入元素外包 `<w:ins>`，并自动用 `next_revision_id` 避开原文档已有修订 id（全文档 `w:id` 须唯一）；`track_changes=False` 时退化为普通插入。

## 已知限制

1. 仅支持 `.docx`（OOXML），**不支持旧版 `.doc`**（需 Word COM/pywin32，本环境未装）。
2. 查找替换不覆盖**页眉/页脚**。
3. 读取时 Word 列表样式按普通段落处理，不还原编号/项目符号；往返后列表符号会丢失。
4. 合并单元格的表格，`row.cells` 可能返回重复单元格引用。
5. 在 Windows 的 GBK 终端（cmd）中直接查看中文输出会乱码，请先执行 `chcp 65001`，或使用 `-o` 输出到文件后查看。

## 开发约定

- 新增能力优先在 `docwork/` 包内实现，再在 `cli.py` 暴露子命令、在 `scripts/` 加便捷包装。
- 所有读写文件统一使用 UTF-8 编码（`open(..., encoding="utf-8")`）。
- 修改代码后，用 `$PY -m docwork ...` 跑一遍冒烟测试（参考 `examples/sample.md` 的往返）。
