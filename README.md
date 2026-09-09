# docwork

对 Word 文档（`.docx`）进行**读取 / 修改 / 撰写**的轻量工具集，仅依赖 `python-docx`。

## 环境

使用 Anaconda 的 `docx` 环境：

```bash
D:/anaconda3/envs/docx/python.exe --version   # Python 3.12.14
```

## 快速开始

```bash
PY=D:/anaconda3/envs/docx/python.exe

# 1. 从 Markdown 撰写文档
$PY -m docwork write examples/sample.md -o docs/output/sample.docx

# 2. 读回为 Markdown / 纯文本 / JSON
$PY -m docwork read docs/output/sample.docx -f markdown
$PY -m docwork read docs/output/sample.docx -f json -o out.json

# 3. 修改（查找替换、增删段落）
$PY -m docwork edit docs/output/sample.docx -o docs/output/edited.docx \
    --replace 张三=李四 --append "追加段落" --prepend "前插段落"

# 4. 修订模式合并（按 JSON 计划插入新内容，可在 Word 审阅视图中接受/拒绝）
$PY -m docwork merge docs/plans/third_party_merge.json \
    --src docs/input/源.docx --dst docs/output/输出.docx
```

## 四种操作

| 操作 | 子命令 | 说明 |
| --- | --- | --- |
| 读取 | `read` | `.docx` → Markdown / 纯文本 / 结构化 JSON |
| 撰写 | `write` | Markdown / JSON → `.docx` |
| 修改 | `edit` | 查找替换、追加 / 前插 / 删除段落 |
| 修订合并 | `merge` | 按 JSON 计划以修订模式（Track Changes）插入内容 |

完整的命令参数与代码调用示例见 [CLAUDE.md](CLAUDE.md)。
