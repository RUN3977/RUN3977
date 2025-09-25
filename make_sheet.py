# -*- coding: utf-8 -*-
"""
英语单词练习纸生成脚本
----------------------
- 页眉仅包含 Date / Section（无页脚）
- 表头+交替行配色，可自定义更多颜色
- 自动分页，可控制行数、行高与列宽
- 所有关键参数均可通过命令行调节（带中文注释）

依赖：
    pip install reportlab

示例：
    python make_sheet.py \
        --output word_sheet.pdf \
        --rows 40 --rows-per-page 18 \
        --page-size A4 \
        --margins 18,18,36,36 \
        --col-widths 32,120,90,160,130 \
        --header-bg lightblue \
        --row-colors lightyellow,whitesmoke \
        --grid-color grey \
        --header-text-color black \
        --body-text-color darkslategrey \
        --row-height 28 --header-row-height 32 \
        --title "Word Practice" \
        --title-size 18 \
        --section "CS-101" \
        --date "2025-09-24" \
        --no-auto-number
"""

from __future__ import annotations

import argparse
import os
import sys
from typing import List, Optional, Sequence, Tuple

import shlex

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, LETTER, landscape, portrait
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


# --------------------------------
# 供 PyCharm 直接运行时的示例参数
# --------------------------------
# 当在 PyCharm 中点击 “Run” 且没有配置额外参数时，PyCharm 会设置
# ``PYCHARM_HOSTED`` 环境变量。此时脚本会自动使用下方的默认参数，
# 方便初学者无需额外配置即可生成示例 PDF。可根据需要自行修改。
PYCHARM_DEFAULT_ARGS: Tuple[str, ...] = (
    "--output",
    "pycharm_word_sheet.pdf",
    "--rows",
    "36",
    "--rows-per-page",
    "18",
    "--title",
    "Word Practice",
    "--section",
    "CS-101",
    "--date",
    "2025-09-24",
)


# ---------------------------
# 工具：颜色解析（支持命名色 / #RRGGBB）
# ---------------------------
def parse_color(value: Optional[str]):
    """将颜色字符串解析为 reportlab ``Color``。

    支持命名颜色（如 ``lightblue``）或 ``#RRGGBB`` 形式的 16 进制颜色值。
    传入 ``None`` 时返回 ``None``。
    """

    if value is None:
        return None
    value = value.strip()
    if not value:
        return None

    if value.startswith("#") and len(value) == 7:
        r = int(value[1:3], 16) / 255.0
        g = int(value[3:5], 16) / 255.0
        b = int(value[5:7], 16) / 255.0
        return colors.Color(r, g, b)

    color_obj = getattr(colors, value, None)
    if color_obj is None:
        raise ValueError(f"Unknown color: {value}")
    return color_obj


# ---------------------------
# 主函数：生成 PDF
# ---------------------------
def generate_word_practice_pdf(
    output_path: str,
    # ——版式参数——
    page_size: str = "A4",                          # 页面尺寸：A4 / A4-L / LETTER / LETTER-L
    margins: Tuple[float, float, float, float] = (18, 18, 36, 36),
    # (left, right, top, bottom) 页面边距（单位：pt；72pt=1英寸；常用 18=0.25in、36=0.5in）
    title: Optional[str] = None,                    # 可选标题（置于页眉下方的正文上方）
    title_size: int = 18,                           # 标题字号
    title_color: str = "black",                    # 标题颜色
    # ——表格参数——
    rows: int = 30,                                 # 总行数（跨页自动分页）
    rows_per_page: int = 15,                        # 每页行数
    header_labels: Optional[Sequence[str]] = None,  # 表头文字
    col_widths: Sequence[float] = (32, 120, 90, 160, 130),  # 各列宽（pt）
    show_row_numbers: bool = True,                  # 是否显示首列编号(No.)
    auto_number: bool = True,                       # 若显示 No. 列，是否自动填充编号
    header_row_height: Optional[float] = None,      # 表头行高（pt）
    row_height: Optional[float] = None,             # 正文行高（pt）
    # ——配色与字体——
    header_bg: str = "lightblue",                   # 表头背景色
    row_colors: Sequence[str] = ("lightyellow", "whitesmoke"),  # 交替行背景色
    grid_color: str = "grey",                       # 表格网格线颜色
    header_font: str = "Helvetica-Bold",            # 表头字体
    body_font: str = "Helvetica",                   # 正文字体
    header_font_size: int = 12,                     # 表头字号
    body_font_size: int = 11,                       # 正文字号
    header_text_color: str = "black",               # 表头文字颜色
    body_text_color: str = "black",                 # 正文文字颜色
    # ——页眉文本（仅 Date / Section，不生成页脚）——
    date_text: Optional[str] = None,                # 页眉右侧 Date 文本（为空则输出空线）
    section_text: Optional[str] = None,             # 页眉左侧 Section 文本（为空则输出空线）
):
    """生成“英语单词练习纸”PDF（仅页眉，包含 Date / Section；无页脚）。"""

    # 默认表头
    default_headers = ["No.", "Word", "Phonetic", "Meaning", "Spelling Practice"]
    header_labels = list(header_labels) if header_labels is not None else default_headers

    # 若不显示编号列，且首列为 No.，则移除首列
    effective_headers = list(header_labels)
    effective_col_widths = list(col_widths)
    if not show_row_numbers and effective_headers:
        first_label = effective_headers[0].lower().replace("。", ".")
        if first_label.startswith("no"):
            effective_headers = effective_headers[1:]
            if len(effective_col_widths) >= len(header_labels):
                effective_col_widths = effective_col_widths[1:]

    if len(effective_headers) != len(effective_col_widths):
        raise ValueError("col_widths 数量必须与表头列数一致。")

    # 页面尺寸
    ps = page_size.upper().strip()
    if ps == "A4":
        pagesize = portrait(A4)
    elif ps == "A4-L":
        pagesize = landscape(A4)
    elif ps == "LETTER":
        pagesize = portrait(LETTER)
    elif ps == "LETTER-L":
        pagesize = landscape(LETTER)
    else:
        raise ValueError("page_size must be one of: A4, A4-L, LETTER, LETTER-L")

    # 颜色
    header_bg_c = parse_color(header_bg)
    grid_c = parse_color(grid_color)
    header_text_c = parse_color(header_text_color) or colors.black
    body_text_c = parse_color(body_text_color) or colors.black
    row_cs = [parse_color(c) for c in row_colors] if row_colors else [colors.whitesmoke]

    # 文档与页边距
    left, right, top, bottom = margins
    doc = SimpleDocTemplate(
        output_path,
        pagesize=pagesize,
        leftMargin=left,
        rightMargin=right,
        topMargin=top,
        bottomMargin=bottom,
    )
    styles = getSampleStyleSheet()
    story = []

    # 标题（可选）
    if title:
        title_style = styles["Title"].clone("WordPracticeTitle")
        title_style.fontName = header_font
        title_style.fontSize = title_size
        title_style.textColor = parse_color(title_color) or colors.black
        story.append(Paragraph(title, title_style))
        story.append(Spacer(1, 8))

    # ——页眉绘制：仅 Section / Date，无页脚——
    def draw_header(canv: canvas.Canvas, _doc):
        width, height = pagesize
        canv.setFont("Helvetica", 10)

        # 左：Section
        if section_text:
            canv.drawString(_doc.leftMargin, height - 20, f"Section: {section_text}")
        else:
            canv.drawString(_doc.leftMargin, height - 20, "Section: ______________________")

        # 右：Date
        if date_text:
            canv.drawRightString(width - _doc.rightMargin, height - 20, f"Date: {date_text}")
        else:
            canv.drawRightString(width - _doc.rightMargin, height - 20, "Date: ______________________")

    # ——分页生成表格——
    total_rows = max(0, rows)
    idx = 0
    num_cols = len(effective_headers)

    if num_cols == 0:
        raise ValueError("表头不能为空。请至少保留一列。")

    while idx < total_rows or (total_rows == 0 and idx == 0):
        take = min(rows_per_page, total_rows - idx) if total_rows > 0 else 0

        # 表头 + 数据行
        table_data: List[List[str]] = [list(effective_headers)]
        for r in range(take):
            row_cells = ["" for _ in range(num_cols)]
            if show_row_numbers and auto_number:
                row_cells[0] = str(idx + r + 1)
            table_data.append(row_cells)

        # 若无数据行，也至少保留空白行以便打印
        if take == 0:
            row_cells = ["" for _ in range(num_cols)]
            if show_row_numbers and auto_number:
                row_cells[0] = "1"
            table_data.append(row_cells)

        # 表格行高
        row_heights = None
        if header_row_height is not None or row_height is not None:
            header_h = header_row_height if header_row_height is not None else row_height
            body_h = row_height if row_height is not None else header_h
            row_heights = [header_h] + [body_h] * (len(table_data) - 1)

        tbl = Table(table_data, colWidths=list(effective_col_widths), rowHeights=row_heights)

        # 表格样式
        ts = TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, grid_c),
            ("BACKGROUND", (0, 0), (-1, 0), header_bg_c),
            ("FONTNAME", (0, 0), (-1, 0), header_font),
            ("FONTSIZE", (0, 0), (-1, 0), header_font_size),
            ("TEXTCOLOR", (0, 0), (-1, 0), header_text_c),
            ("FONTNAME", (0, 1), (-1, -1), body_font),
            ("FONTSIZE", (0, 1), (-1, -1), body_font_size),
            ("TEXTCOLOR", (0, 1), (-1, -1), body_text_c),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ])

        # 交替行背景
        for r in range(1, len(table_data)):
            color = row_cs[(r - 1) % len(row_cs)]
            ts.add("BACKGROUND", (0, r), (-1, r), color)

        tbl.setStyle(ts)
        story.append(tbl)
        idx += take

        if idx < total_rows:
            story.append(PageBreak())

    # 仅挂载页眉（第一页与后续页）
    doc.build(story, onFirstPage=draw_header, onLaterPages=draw_header)
    print(f"✅ PDF saved: {output_path}")


# ---------------------------
# CLI
# ---------------------------
def _parse_float_list(values: str) -> List[float]:
    return [float(x.strip()) for x in values.split(",") if x.strip()]


def _parse_str_list(values: Optional[str]) -> List[str]:
    if values is None:
        return []
    return [x.strip() for x in values.split(",") if x.strip()]


def main(argv: Optional[Sequence[str]] = None):
    """命令行入口，可在 PyCharm/终端中运行。

    ``argv`` 可传入自定义参数列表，默认读取 ``sys.argv[1:]``。
    若在 PyCharm 中直接点击运行且未设置参数（即 ``argv`` 为空），脚本会
    自动使用 ``PYCHARM_DEFAULT_ARGS``，方便快速生成示例 PDF。
    也可以传入字符串，本函数会自动拆分为参数列表。
    """

    if argv is None:
        argv = sys.argv[1:]

    if isinstance(argv, str):
        argv = shlex.split(argv)

    if not argv and os.environ.get("PYCHARM_HOSTED"):
        argv = list(PYCHARM_DEFAULT_ARGS)

    ap = argparse.ArgumentParser(
        description="Generate English Word Practice PDF (Header only: Date / Section)."
    )
    # 输出文件路径
    ap.add_argument("--output", type=str, default="english_word_practice_colored.pdf", help="输出 PDF 文件路径")
    # 行数控制
    ap.add_argument("--rows", type=int, default=30, help="总行数（自动分页）")
    ap.add_argument("--rows-per-page", type=int, default=25, help="每页行数")
    # 页面与边距
    ap.add_argument(
        "--page-size",
        type=str,
        default="A4",
        choices=["A4", "A4-L", "LETTER", "LETTER-L"],
        help="页面尺寸：A4 / A4-L（横向）/ LETTER / LETTER-L（横向）",
    )
    ap.add_argument(
        "--margins",
        type=str,
        default="18,18,36,36",
        help="四周边距 L,R,T,B（单位pt；72pt=1英寸），示例：18,18,36,36",
    )
    # 表格与样式
    ap.add_argument("--col-widths", type=str, default="32,120,90,160,130", help="各列宽度（pt），逗号分隔")
    ap.add_argument("--header-bg", type=str, default="lightblue", help="表头背景色（命名色或#RRGGBB）")
    ap.add_argument(
        "--row-colors",
        type=str,
        default="lightyellow,whitesmoke",
        help="交替行背景色列表，逗号分隔，循环使用",
    )
    ap.add_argument("--grid-color", type=str, default="grey", help="表格网格线颜色")
    ap.add_argument("--header-font", type=str, default="Helvetica-Bold", help="表头字体")
    ap.add_argument("--body-font", type=str, default="Helvetica", help="正文字体")
    ap.add_argument("--header-font-size", type=int, default=12, help="表头字号")
    ap.add_argument("--body-font-size", type=int, default=11, help="正文字号")
    ap.add_argument("--header-text-color", type=str, default="black", help="表头文字颜色")
    ap.add_argument("--body-text-color", type=str, default="black", help="正文文字颜色")
    ap.add_argument("--row-height", type=float, default=None, help="正文行高（pt），为空则使用默认自适应")
    ap.add_argument(
        "--header-row-height",
        type=float,
        default=None,
        help="表头行高（pt），为空则与正文一致",
    )
    # 标题
    ap.add_argument("--title", type=str, default=None, help="可选标题文本")
    ap.add_argument("--title-size", type=int, default=18, help="标题字号")
    ap.add_argument("--title-color", type=str, default="black", help="标题颜色")
    # 是否显示编号列
    ap.add_argument("--no-row-numbers", action="store_true", help="不显示 No. 编号列")
    ap.add_argument(
        "--no-auto-number",
        action="store_true",
        help="显示 No. 列但不自动填写编号（保留空白供手写）",
    )
    # 自定义表头（如需改列名）
    ap.add_argument(
        "--header-labels",
        type=str,
        default=None,
        help="自定义表头（逗号分隔）。默认：No.,Word,Phonetic,Meaning,Spelling Practice",
    )
    # 页眉文本
    ap.add_argument("--date", type=str, default=None, help="页眉 Date 文本（为空则打印空线）")
    ap.add_argument("--section", type=str, default=None, help="页眉 Section 文本（为空则打印空线）")

    args = ap.parse_args(argv)

    margins = _parse_float_list(args.margins)
    if len(margins) != 4:
        raise ValueError("margins 需包含 4 个数字：L,R,T,B")

    col_widths = _parse_float_list(args.col_widths)
    header_labels = _parse_str_list(args.header_labels) if args.header_labels else None
    row_colors = _parse_str_list(args.row_colors)

    generate_word_practice_pdf(
        output_path=args.output,  # 输出 PDF 文件路径
        page_size=args.page_size,  # 页面尺寸: "A4" / "A4-L" / "LETTER" / "LETTER-L"
        margins=tuple(margins),  # 页面边距(左,右,上,下) 单位pt; 72pt=1英寸
        title=args.title,  # 可选: 页眉下方的标题文本
        title_size=args.title_size,  # 标题字号
        title_color=args.title_color,  # 标题颜色
        rows=args.rows,  # 总行数（自动分页）
        rows_per_page=args.rows_per_page,  # 每页行数，超出自动分页
        header_labels=header_labels,  # 自定义表头
        col_widths=col_widths,  # 各列宽度 (pt)
        show_row_numbers=not args.no_row_numbers,  # 是否显示 No. 编号列
        auto_number=not args.no_auto_number,  # 是否自动填写编号
        header_row_height=args.header_row_height,  # 表头行高
        row_height=args.row_height,  # 正文行高
        header_bg=args.header_bg,  # 表头背景色
        row_colors=row_colors,  # 交替行背景色，可多个，循环使用
        grid_color=args.grid_color,  # 网格线颜色
        header_font=args.header_font,  # 表头字体
        body_font=args.body_font,  # 正文字体
        header_font_size=args.header_font_size,  # 表头字号
        body_font_size=args.body_font_size,  # 正文字号
        header_text_color=args.header_text_color,  # 表头文字颜色
        body_text_color=args.body_text_color,  # 正文文字颜色
        date_text=args.date,  # 页眉右侧 Date 文本 (为空则打印空线)
        section_text=args.section,  # 页眉左侧 Section 文本 (为空则打印空线)
    )


if __name__ == "__main__":
    main()
