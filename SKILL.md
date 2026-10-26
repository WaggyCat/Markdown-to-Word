---
name: md-to-docx
description: 把 Markdown (.md) 转成 Word (.docx) 并套用固定中文论文/公文排版（宋体 + Times New Roman、正文小四、标题四号且分多级、全文两端对齐、段前段后 0、首行缩进 2 字符、1.5 倍行距、Unicode 上下标转 Word 上下标、删中英文空格）；也能反向检查已有 .docx 里残留的 Markdown 标记（`**加粗**`、`## 标题`、`- 列表`、`| 表格 |`、`[文本](链接)`、`&nbsp;`）并输出清理后的新文件和检查报告。Use this whenever the user wants a .md file exported/converted to Word, asks for 论文/报告/公文排版（宋体、新罗马、小四、标题四号、两端对齐、首行缩进2字符、1.5倍行距、上下标、中英文空格）, or asks to check/clean a Word file that has leftover markdown syntax, even if they don't name a skill.
license: MIT
metadata:
  version: "1.2.0"
  requires: python-docx
---

# Markdown → Word（固定中文排版）

把一份 `.md` 转成 `.docx`，格式按中文论文/公文习惯固定下来。转换是**纯本地**的：
`scripts/md2docx.py` 读 Markdown，用 python-docx 直接写 OOXML，不依赖 Word/WPS 或联网。
反向的检查/清理用 `scripts/mdclean.py`（见“检查与修复”）。

## 什么时候用

- 用户要把 Markdown 变成 Word：`.md 转 word`、`导出 docx`、`转成 Word 发给我`。
- 用户在描述排版要求：宋体 + 新罗马、小四、首行缩进 2 字符、1.5 倍行距、段前段后 0、公式上下标。
- 用户有一篇 Markdown 稿件要交给需要 Word 的人（投稿、周报、公文）。
- **反过来**：用户拿来一份 .docx，想查里面有没有 `**`、`##`、`- ` 这类 Markdown 残留，
  或说“这份 Word 排版乱了/有星号”，→ 用 `scripts/mdclean.py`（见下面“检查与修复”）。

只需要**读**文档内容来回答问题时不要用这个技能——那种情况走 `mcp__cherry-tools__to_markdown`。

## 跑起来

工作目录随便，脚本路径用绝对路径（下面的 `<SKILL>` 指本 SKILL.md 所在目录）：

```bash
uv run --with python-docx python "<SKILL>/scripts/md2docx.py" "输入.md" -o "输出.docx"
```

- 不给 `-o` 时输出同目录同名 `.docx`；可以一次传多个 md 批量转换（此时不能用 `-o`）。
- 机器上已有 python + python-docx 时，直接 `python md2docx.py 输入.md` 也行；`uv` 是 Cherry 自带的，无需预装。
- 转换成功后可以用 `mcp__cherry-tools__report_artifacts` 把 docx 作为交付物报给用户。

## 检查与修复：已有 Word 里的 Markdown 残留

两个脚本，方向相反：

| 脚本 | 方向 |
|---|---|
| `scripts/md2docx.py` | `.md` → 新 `.docx` |
| `scripts/mdclean.py` | 已有 `.docx` → 查残留 → 清理后的新 `.docx` + `.md` 检查报告 |

```bash
uv run --with python-docx python "<SKILL>/scripts/mdclean.py" "输入.docx"
# → 输入.cleaned.docx  + 输入.cleaned.cleanreport.md

uv run --with python-docx python "<SKILL>/scripts/mdclean.py" "输入.docx" -o "输出.docx" --report "报告.md"
uv run --with python-docx python "<SKILL>/scripts/mdclean.py" "输入.docx" --dry-run   # 只检查，不产出 docx
```

它逐段（正文、表格单元格、页眉页脚）检查 16 类残留，命中就改并记录：

`**加粗**` / `__加粗__`（真加粗）、`*斜体*` / `_斜体_`（真斜体）、`~~删除线~~`、`` `行内代码` ``（转等宽）、
`[文本](链接)`、`![图片](src)`（留显示文本）、`# 标题`、`> 引用`、`- ` / `* ` / `+ `、`1. `、`---`、
`| a | b |` 表格行、`|---|` 表格分隔行、`<br>` 等 HTML 标签、`&nbsp;` 等实体、`\*` 反斜杠转义、
以及 Unicode 上下标。

报告（`*.cleanreport.md`）分三块：**命中汇总**（每类各多少处）、**明细**（位置 + 原文片段 + 处理动作）、
**未处理 / 需人工确认**（跨 run 或有图片/换行/域而没能安全处理的）。回答用户时先给汇总数字，
再把“未处理”列给他——那几处必须人工看。

可选开关：

```bash
--dry-run              只出报告，不改文件
--keep-list-markers    不删行首 `-` / `1.` / `>`（它们其实是正文编号时）
--no-convert-scripts   不转换 Unicode 上下标
--remove-cjk-space     顺便删中英文之间的空格（默认不动）
--fix-headings         把 `#` 级别写回段落大纲级别（便于生成目录）
--reformat             清理后再套用本技能的标准排版（宋体/新罗马、两端对齐、首行缩进……）
```

安全边界（跟用户讲清楚，别让他误以为全自动重排）：

- **源文件永不被改**；`-o` 不许指向源文件。
- 跨 run 断开的标记（如 `**` 被工具切成两半）能正确合并处理；重建时**复制原 run 的 rPr**，
  所以字体/字号/颜色不丢。
- 含图片、硬换行、域的段落走**保守模式**：只删标记文本，不套加粗/斜体，并在报告里标注。
- 列表标记删掉后**不会**自动变成 Word 列表；`| a | b |` 只去竖线、**不会**重建成真表格。
- 行首 `1. ` 这类编号与 Markdown 有序列表在文本上无法区分——默认按列表处理，遇到论文编号让用户加 `--keep-list-markers`。
- 同一份文件**只跑一次**：`\*` 清理成字面量 `*` 后再跑会被当成斜体标记。

## 默认格式（= 用户没额外要求时就用这套）

正文部分：

| 项 | 值 | 实现 |
|---|---|---|
| 中文 / 西文 | 宋体 / Times New Roman | `w:rFonts` 的 `eastAsia` + `ascii/hAnsi`，并剥离主题字体 |
| 正文字号 | 小四 12pt | `w:sz=24` |
| 对齐 | **全部两端对齐** | `w:jc="both"`（代码块例外，左对齐） |
| 段前 / 段后 | 0 / 0 | `w:spacing` |
| 首行缩进 | 2 字符 | `w:ind firstLineChars="200"`（真·字符单位，随字号自适应） |
| 行距 | 1.5 倍 | `w:line=360 lineRule=auto` |
| 上下标 | Unicode → Word 上下标 | 逐字符映射后写 `w:vertAlign`，连续上下标合并成一个 run |
| 中英文之间空格 | 删除 | 例如 `中文 English` → `中文English`；代码块内不动 |

多级标题：**各级一律四号 (14pt)、段前段后 0、1.5 倍行距**，靠对齐与缩进分级：

| 级别 | 对齐 | 加粗 | 首行缩进 | 大纲级别 |
|---|---|---|---|---|
| `#` 一级 | 居中 | ✅ | 无 | 1 |
| `##` 二级 | 两端对齐 | ✅ | 无（顶格） | 2 |
| `###`～`######` 三级及以下 | 两端对齐 | ✅ | 2 字符 | 3～6 |

即使标题有级别差异，字号仍然统一，所以层次的区分完全靠“居中 / 顶格 / 缩进”与加粗——
这是中文公文里最不容易出错的组合。要改（比如一级也两端对齐、全部不加粗），
改脚本顶部 `HEADING_LAYOUT` 表，或用 `--heading-justify` / `--no-heading-bold`。

其他：代码块 Consolas、左对齐、不缩进；表格 Table Grid 边框、单元格两端对齐；
列表保留 `•` / `1.` 前缀（不依赖 Word 编号域）。

## 需要改格式时

先看用户到底要改什么，再用参数覆盖，**不要**去改脚本里的常量：

```bash
--cn-font 黑体  --en-font Arial        # 换字体
--size 14                              # 正文字号（pt），14 = 四号
--heading-size 16                      # 标题字号，16 = 三号（默认 14 = 四号）
--align left                           # 正文对齐：justify(默认) / left / center / right
--line-spacing 2                       # 行距倍数；也可 1.25
--first-line-indent 0                  # 正文首行缩进字符数；0 = 不缩进
--space-before 6 --space-after 6        # 段前段后（pt）
--heading-justify                      # 一级标题也两端对齐（默认居中）
--no-heading-bold                      # 标题全部不加粗
--indent-code / --indent-table          # 代码块 / 表格内文字也首行缩进
--keep-cjk-spacing                     # 保留中英文之间的空格（默认删除）
```

用户说"段后 1.5 行"这类非倍数需求时，用 `--space-after`（pt）换算，别指望 `--line-spacing`。

标题序号：脚本**不自动编号**（Markdown 里写了 `## 1.2 方法` 就原样保留）。用户要“自动编号 1 / 1.1 / 1.1.1”时，
那是 Word 的多级列表（`w:numPr`），脚本没实现——先说明这一点，别假装支持。

## 容易踩的坑

- **不要改源文件**。脚本只读 `.md`，只写新 `.docx`；输出的路径已存在时会覆盖，拿不准就先问用户。
- **脚本输出的是 UTF-8**，Windows 控制台是 GBK 时不要去掉 `main()` 里的 `sys.stdout.reconfigure`，否则中文和符号会抛 `UnicodeEncodeError`。
- **`w:pPr` / `w:rPr` 子元素顺序是硬约束**，Word 对乱序会报"内容不可读"。脚本用 `PPR_SEQ` / `RPR_SEQ` 有序插入，新增属性时务必走 `get_or_add_child()` / `_ordered_child()`，不要直接 `append`。
- 只支持 Markdown 的子集（标题、段落、硬换行、粗斜体/删除线/行内代码/链接、列表、引用、围栏代码块、GFM 表格）。遇到 `<div>`、脚注、数学公式 `$...$` 原样当普通文本输出——公式场景要先跟用户确认，别自作主张。
- 中英文空格规则：只在"中文↔拉丁/数字"交界处删空格；中文之间、英文之间的空格保留，超/下标字符算拉丁侧。

## 交付前自检

转换后用 python-docx 抽查一段，确认字体、缩进、行距、上下标都落到了 XML 上（这是唯一可信的验证方式，看文件大小说明不了问题）：

```bash
uv run --with python-docx python -c "
import docx; from docx.oxml.ns import qn
d = docx.Document('输出.docx')
for p in d.paragraphs[:6]:
    if not p.text.strip(): continue
    pPr = p._p.pPr; r = p.runs[0]._element.rPr
    print(p.text[:12], 'jc=', pPr.find(qn('w:jc')).get(qn('w:val')),
          'sz=', r.find(qn('w:sz')).get(qn('w:val')),
          'line=', pPr.find(qn('w:spacing')).get(qn('w:line')))
"
```

期望：正文 `jc=both sz=24`，标题 `jc=both/center sz=28`，都 `line=360`。
