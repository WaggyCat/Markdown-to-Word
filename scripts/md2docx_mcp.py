#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2docx MCP server（stdio）—— 把 md2docx.py 暴露成结构化工具 `md_to_docx`。

启动（Cherry 注册时同样用这条命令，路径换成你自己的绝对路径）：
    uv run --with python-docx --with "mcp[cli]" python "C:/.../scripts/md2docx_mcp.py"

工具：md_to_docx(md_path, docx_path?, ...格式参数)
"""

from __future__ import annotations

import sys
from pathlib import Path

# 让本文件能 import 同目录的 md2docx.py
sys.path.insert(0, str(Path(__file__).resolve().parent))

import md2docx as core  # noqa: E402

# MCP SDK 1.x 的类叫 FastMCP，2.x 重命名为 MCPServer；两个都兼容
try:  # mcp >= 2.0
    from mcp.server.mcpserver import MCPServer as _Server
    _SDK = "2"
except ImportError:  # mcp 1.x
    from mcp.server.fastmcp import FastMCP as _Server  # type: ignore[no-redef]
    _SDK = "1"

mcp = _Server("md-to-docx")


@mcp.tool()
def md_to_docx(
    md_path: str,
    docx_path: str | None = None,
    cn_font: str = "宋体",
    en_font: str = "Times New Roman",
    font_size: float = 12,
    first_line_indent_chars: float = 2,
    line_spacing: float = 1.5,
    space_before: float = 0,
    space_after: float = 0,
    body_align: str = "justify",
    heading_size: float = 14,
    heading_bold: bool = True,
    heading_center_l1: bool = True,
    remove_cjk_latin_space: bool = True,
    indent_code: bool = False,
    indent_table: bool = False,
) -> str:
    """把 Markdown (.md) 转成 Word (.docx)，套用固定中文论文/公文排版。

    默认格式：中文宋体 + 英文 Times New Roman；正文小四(12pt)、标题四号(14pt)；
    全部两端对齐；段前段后 0；首行缩进 2 字符；1.5 倍行距；Unicode 上下标
    转为 Word 上下标；删除中英文之间的空格。源 .md 只读，不改动。

    Args:
        md_path: 输入 Markdown 文件的绝对路径。
        docx_path: 输出 .docx 路径；省略则与输入同目录同名。
        cn_font: 中文字体，默认 宋体。
        en_font: 西文字体，默认 Times New Roman。
        font_size: 正文字号（pt），默认 12 = 小四。
        first_line_indent_chars: 正文首行缩进字符数，默认 2；0 = 不缩进。
        line_spacing: 行距倍数，默认 1.5。
        space_before: 段前（pt），默认 0。
        space_after: 段后（pt），默认 0。
        body_align: 正文对齐，justify（两端对齐，默认）/ left / center / right。
        heading_size: 各级标题字号（pt），默认 14 = 四号。
        heading_bold: 标题是否加粗，默认 True。
        heading_center_l1: 一级标题是否居中（默认 True）；False = 也两端对齐。
        remove_cjk_latin_space: 是否删除中文与英文/数字之间的空格，默认 True。
        indent_code: 代码块是否也首行缩进，默认 False。
        indent_table: 表格内文字是否也首行缩进，默认 False。

    Returns:
        一行结果说明，包含输出文件的绝对路径。
    """
    core.CN_FONT = cn_font
    core.EN_FONT = en_font
    core.FONT_SIZE = float(font_size)
    core.FIRST_LINE_CHARS = int(round(float(first_line_indent_chars) * 100))
    core.LINE_SPACING = float(line_spacing)
    core.SPACE_BEFORE = float(space_before)
    core.SPACE_AFTER = float(space_after)
    core.BODY_ALIGN = body_align
    core.HEADING_SIZE = float(heading_size)
    core.HEADING_BOLD = bool(heading_bold)
    core.HEADING_CENTER_L1 = bool(heading_center_l1)
    core.REMOVE_CJK_LATIN_SPACE = bool(remove_cjk_latin_space)
    core.CODE_NO_INDENT = not indent_code
    core.TABLE_NO_INDENT = not indent_table

    src = Path(md_path).expanduser()
    if not src.is_file():
        return f"错误：找不到输入文件 {src}"

    out = core.convert(src, Path(docx_path).expanduser() if docx_path else None)
    return f"已生成 Word 文件：{out.resolve()}"


if __name__ == "__main__":
    print(f"[md-to-docx] MCP server 启动（SDK {_SDK}，stdio）", file=sys.stderr, flush=True)
    mcp.run()  # 默认 stdio
