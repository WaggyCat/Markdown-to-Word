# md-to-docx

把 Markdown 转成排版规整的中文 Word 文档，并反向检查已有 Word 里残留的 Markdown 标记。
纯本地运行，不依赖 Word / WPS / 联网。

A local, dependency-light tool to convert Markdown into a properly typeset Chinese
`.docx` — and to clean leftover Markdown syntax out of existing Word files.

## 它能做什么

**1. Markdown → Word**（`scripts/md2docx.py`）

固定套用中文论文/公文格式：

| 项 | 默认值 |
|---|---|
| 中文字体 / 西文字体 | 宋体 / Times New Roman |
| 正文字号 | 小四（12 pt） |
| 标题 | 各级一律四号（14 pt）、加粗，按「居中 / 顶格 / 缩进 2 字符」分级 |
| 对齐 | 全文两端对齐（代码块除外） |
| 段前 / 段后 | 0 / 0 |
| 首行缩进 | 2 字符（真·字符单位，随字号自适应） |
| 行距 | 1.5 倍 |
| Unicode 上下标 | `² ³ ₁ ⁻ ⁿ` → Word 真正的上标 / 下标 |
| 中英文之间的空格 | 自动删除（`中文 English` → `中文English`） |

支持标题（保留大纲级别，可一键生成目录）、段落、硬换行、粗体/斜体/删除线/行内代码/链接、
有序与无序列表、引用、围栏代码块、GFM 表格。

**2. Word → 检查 Markdown 残留**（`scripts/mdclean.py`）

扫描正文、表格单元格、页眉页脚，找出 16 类残留并清理，输出新文件 + Markdown 检查报告：

```
**加粗**  __加粗__  *斜体*  _斜体_  ~~删除线~~  `行内代码`
[文本](链接)  ![图](src)  # 标题  > 引用  - / * / + 列表  1. 列表
--- 分隔线  | a | b | 表格行  |---| 表格分隔行  <br>  &nbsp;  \*  ² ³ ₁
```

报告包含命中汇总、逐处明细（位置 / 原文片段 / 处理动作）以及「未处理、需人工确认」清单。

## 安装

### Cherry Studio

把整个文件夹放进 Cherry 的托管技能目录，重启后即可在技能列表看到 `md-to-docx`：

| 平台 | 目录 |
|---|---|
| Windows | `%APPDATA%\CherryStudio\Data\Skills\` |
| macOS | `~/Library/Application Support/CherryStudio/Data/Skills/` |
| Linux | `~/.config/CherryStudio/Data/Skills/` |

```bash
# 或者直接 clone 进去（注意：目标目录名必须是 md-to-docx，和 SKILL.md 里的 name 一致）
cd /c/Users/<you>/AppData/Roaming/CherryStudio/Data/Skills
git clone https://github.com/WaggyCat/Markdown-to-Word.git md-to-docx
```

### Claude Code / 其他支持 Agent Skills 的环境

把目录放到对应的 skills 目录（例如 `~/.claude/skills/md-to-docx/`）即可。

### 直接用脚本（不需要任何 skill 机制）

```bash
pip install python-docx
# 或用 uv（无需预装 Python 依赖）
uv run --with python-docx python scripts/md2docx.py 输入.md -o 输出.docx
```

## 用法

```bash
# Markdown → Word
uv run --with python-docx python scripts/md2docx.py 输入.md
uv run --with python-docx python scripts/md2docx.py 输入.md -o 输出.docx
uv run --with python-docx python scripts/md2docx.py a.md b.md          # 批量

# 常用覆盖参数
--cn-font 黑体 --en-font Arial     # 换字体
--size 14                          # 正文字号（pt）
--heading-size 16                  # 标题字号（16 = 三号）
--line-spacing 2                   # 行距倍数
--first-line-indent 0              # 首行缩进字符数，0 = 不缩进
--space-before 6 --space-after 6   # 段前段后（pt）
--align left                       # 正文对齐：justify(默认)/left/center/right
--heading-justify                  # 一级标题也两端对齐（默认居中）
--keep-cjk-spacing                 # 保留中英文之间的空格
```

```bash
# Word 残留检查与清理
uv run --with python-docx python scripts/mdclean.py 输入.docx
# → 输入.cleaned.docx + 输入.cleaned.cleanreport.md

--dry-run              只出报告，不改文件
--keep-list-markers    行首 - / 1. / > 其实是正文编号时用
--no-convert-scripts   不转换 Unicode 上下标
--remove-cjk-space     顺便删中英文空格（默认不动）
--fix-headings         把 # 级别写回段落大纲级别
--reformat             清理后再套用上面的标准排版
```

## 已知限制

- 格式常量在脚本顶部，或改 `HEADING_LAYOUT` 表自定义各级标题。
- 标题**不自动编号**；需要 `1 / 1.1 / 1.1.1` 请在 Markdown 里写好，或在 Word 里用多级列表。
- `mdclean.py` 删除列表标记后**不会**生成 Word 列表；`| a | b |` 只去竖线、不重建成真表格。
- 含图片、硬换行、域的段落走保守模式：只删文本标记，不套加粗/斜体。
- 同一份文件建议只跑一次：`\*` 清理成字面 `*` 后再跑会被当成斜体标记。
- macOS / Linux 需自备「宋体」字体，否则会回退到别的字体；也可用 `--cn-font` 换成系统里有的字体。
- 只在 `.docx` 上工作；`.doc` 请先另存为 `.docx`。

## 实现要点（想改代码的话）

- 字体用 `w:rFonts` 的 `eastAsia` + `ascii/hAnsi` 分别控制中西文，并剥离主题字体属性。
- 首行缩进写 `w:ind w:firstLineChars="200"`（字符单位），另带 `firstLine` 兜底。
- `w:pPr` / `w:rPr` 的子元素**顺序是硬约束**，乱序会让 Word 报「内容不可读」；
  代码里用 `PPR_SEQ` / `RPR_SEQ` 有序插入，改动时请走 `get_or_add_child()` / `_ordered_child()`。
- `mdclean.py` 重建段落时**复制原 run 的 `rPr`**，因此字体、字号、颜色不会丢；
  跨 run 断开的 `**加粗**` 也能正确合并。

## 可选：作为 MCP 服务器使用

`scripts/md2docx_mcp.py` 把 `md_to_docx` 暴露成 MCP 工具（stdio），兼容 `mcp` SDK 1.x / 2.x：

```json
{
  "name": "md-to-docx",
  "type": "stdio",
  "command": "uv",
  "args": ["run", "--with", "python-docx", "--with", "mcp", "python",
           "/absolute/path/to/scripts/md2docx_mcp.py"]
}
```

## 更新

已 clone 到 Skills 目录的话，拉新版本就是：

```bash
cd /c/Users/<you>/AppData/Roaming/CherryStudio/Data/Skills/md-to-docx
git pull
```

Cherry 会在下次技能同步时重新读取 `SKILL.md`。

## License

MIT
