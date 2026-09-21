# -*- coding: utf-8 -*-
"""
生成「IBGT 部门人员流动看板」工作簿（Excel 原生公式看板，无宏）。

用法:
    python make_hr_dashboard.py           # 生成 模板 + 演示 两个文件

结构:
    看板        KPI 卡片 + 招聘类型筛选下拉框 + 3 张原生图表
    月度明细    48 个月 × (期末在职/当月流出/Perm/PC/MS 流出/当月入职)
    数据_Profile / 数据_Leaver   数据粘贴区（表头已按 IBGT Profile 字段预填）
    说明        使用说明 + 统计口径 + 表头映射配置 + 数据质量自检
    计算        隐藏辅助列（类型规范化、StaffID 去重标记、月份链）

口径:
    期末在职(M)   = StartDate <= 月末 且 (LastWorkDay 为空 或 >= 月末)
    当月流出(M)   = LastWorkDay 落在 M 月
    当月入职(M)   = StartDate 落在 M 月
    Profile 与 Leaver 两表合并统计，同一 StaffID 同时出现在两表时自动去重。
    类型列不区分大小写、去首尾空格后匹配。

依赖: pip install openpyxl
"""

import os
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    Paragraph, ParagraphProperties, CharacterProperties,
    Font as DrawingFont, RichTextProperties,
)

# 样式令牌 base.py 随脚本同目录分发（本地自包含，运行不依赖任何外部安装路径）

"""
xlsx skill — Base Template
===========================
Single source of truth for design tokens, font resolution, and style factories.
All scene/engine code MUST import from here. Never hardcode colors, fonts, or styles.

Usage:
    from templates.base import *

    # To switch palette based on user prompt (call BEFORE creating styles):
    use_palette("帮我做一个温暖的销售月报")  # Chinese prompt example
    # → All color tokens and style factories now use 'warm' palette.

    # Or manually:
    use_palette_explicit("warm")
"""

import platform
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from copy import copy


# ============================================================
# §1  Font Resolution (cross-platform fallback chain)
# ============================================================

def _resolve_font(candidates: list) -> str:
    """Return the first font name likely available on this OS."""
    system = platform.system()
    _platform_hints = {
        "Darwin":  {"Noto Sans SC", "Noto Sans CJK SC", "Source Han Sans SC"},
        "Windows": {"Noto Sans SC", "Noto Sans SC", "Noto Serif SC"},
        "Linux":   {"Noto Sans CJK SC", "WenQuanYi Micro Hei", "Source Han Sans SC"},
    }
    available = _platform_hints.get(system, set())
    for name in candidates:
        if name in available:
            return name
    return candidates[0]


# CJK sans-serif fallback chain
CJK_BODY_CHAIN = [
    "Noto Sans SC",        # Free (OFL), bundled, cross-platform
    "Noto Sans CJK SC",    # Linux / Android
    "Source Han Sans SC",  # Adobe cross-platform (OFL)
    "WenQuanYi Micro Hei", # Linux fallback (GPL)
]

# Latin serif (for formal reports)
LATIN_BODY_CHAIN = [
    "FreeSerif",
    "serif",
]

FONT_CJK   = _resolve_font(CJK_BODY_CHAIN)
FONT_LATIN  = _resolve_font(LATIN_BODY_CHAIN)

# Primary font — CJK font covers ASCII too
FONT_NAME = FONT_CJK

# Bold strategy: heavy-stroke fonts should NOT be bolded
_HEAVY_FONTS = {
    "Noto Sans SC", "Noto Sans CJK SC",
    "Source Han Sans SC", "WenQuanYi Micro Hei",
}
HEADER_BOLD = FONT_NAME not in _HEAVY_FONTS


# ============================================================
# §2  Color Tokens (Three-Color Rule)
# ============================================================

# --- Primary (deep blue — professional default) ---
PRIMARY       = "1B2A4A"
PRIMARY_LIGHT = "D6E4F0"
SECONDARY     = PRIMARY_LIGHT   # derived from primary

# --- Accent (semantic, on-demand) ---
ACCENT_POSITIVE = "1B7D46"      # growth, done, pass   (deep green)
ACCENT_NEGATIVE = "C0392B"      # decline, overdue     (deep red)
ACCENT_WARNING  = "D4820A"      # at-risk, watch       (deep amber)

# --- Neutral (warm gray) ---
NEUTRAL_900 = "37352F"          # body text
NEUTRAL_600 = "8C8A84"          # caption, secondary text
NEUTRAL_200 = "E9E9E8"          # borders, dividers
NEUTRAL_100 = "F7F7F5"          # alternating row fill (odd)
NEUTRAL_50  = "FAFAF9"          # ultra-light bg (optional)
NEUTRAL_0   = "FFFFFF"          # white (even rows)

# --- Header text color (overridable by palette) ---
HEADER_TEXT = "FFFFFF"

# --- Chart palette (max 5 colors) ---
CHART_COLORS = [PRIMARY, ACCENT_POSITIVE, ACCENT_WARNING, ACCENT_NEGATIVE, NEUTRAL_600]

# --- Conditional formatting fills ---
CF_POSITIVE_FILL = PatternFill(bgColor="E8F5E9")
CF_POSITIVE_FONT = Font(color=ACCENT_POSITIVE)
CF_NEGATIVE_FILL = PatternFill(bgColor="FDEDEC")
CF_NEGATIVE_FONT = Font(color=ACCENT_NEGATIVE)
CF_WARNING_FILL  = PatternFill(bgColor="FEF9E7")
CF_WARNING_FONT  = Font(color=ACCENT_WARNING)

# --- Active style (for debugging/logging) ---
_ACTIVE_STYLE = "professional"


# ============================================================
# §2.1  Palette Integration
# ============================================================

def use_palette(prompt: str):
    """
    Auto-detect style from user prompt and switch all color tokens.
    Call this BEFORE creating any styles/cells.

    Three-step matching:
      1. Explicit style keywords → direct match
      2. Scene/content keywords → infer style
      3. No match → professional (safe default)

    Example:
        use_palette("帮我做一个温暖的销售月报")  # Chinese prompt example
        # → 'warm' palette applied
    """
    from templates.palettes import resolve_palette_with_info
    palette, style = resolve_palette_with_info(prompt)
    _apply(palette, style)


def use_palette_explicit(style: str = "professional"):
    """
    Manually select a palette by style name.
    Available: professional, warm, elegant, creative, muji, aesop,
               kinfolk, celine, bottega, chanel, bloomberg, original_blue

    Example:
        use_palette_explicit("warm")
    """
    from templates.palettes import get_palette
    palette = get_palette(style)
    _apply(palette, style)


def get_active_style() -> str:
    """Return the currently active style name."""
    return _ACTIVE_STYLE


def _apply(palette: dict, style: str):
    """Internal: apply a palette dict to all module-level color tokens."""
    global PRIMARY, PRIMARY_LIGHT, SECONDARY
    global ACCENT_POSITIVE, ACCENT_NEGATIVE, ACCENT_WARNING
    global NEUTRAL_900, NEUTRAL_600, NEUTRAL_200, NEUTRAL_100, NEUTRAL_50, NEUTRAL_0
    global CHART_COLORS, HEADER_TEXT
    global CF_POSITIVE_FILL, CF_POSITIVE_FONT
    global CF_NEGATIVE_FILL, CF_NEGATIVE_FONT
    global CF_WARNING_FILL, CF_WARNING_FONT
    global _ACTIVE_STYLE

    PRIMARY       = palette["PRIMARY"]
    PRIMARY_LIGHT = palette["PRIMARY_LIGHT"]
    SECONDARY     = palette["SECONDARY"]
    ACCENT_POSITIVE = palette["ACCENT_POSITIVE"]
    ACCENT_NEGATIVE = palette["ACCENT_NEGATIVE"]
    ACCENT_WARNING  = palette["ACCENT_WARNING"]
    NEUTRAL_900 = palette["NEUTRAL_900"]
    NEUTRAL_600 = palette["NEUTRAL_600"]
    NEUTRAL_200 = palette["NEUTRAL_200"]
    NEUTRAL_100 = palette["NEUTRAL_100"]
    NEUTRAL_50  = palette["NEUTRAL_50"]
    NEUTRAL_0   = palette["NEUTRAL_0"]
    HEADER_TEXT  = palette.get("HEADER_TEXT", "FFFFFF")
    CHART_COLORS = palette["CHART_COLORS"]

    # Rebuild conditional formatting fills/fonts with new accent colors
    CF_POSITIVE_FILL = PatternFill(bgColor=palette.get("CF_POSITIVE_BG", "E8F5E9"))
    CF_POSITIVE_FONT = Font(color=ACCENT_POSITIVE)
    CF_NEGATIVE_FILL = PatternFill(bgColor=palette.get("CF_NEGATIVE_BG", "FDEDEC"))
    CF_NEGATIVE_FONT = Font(color=ACCENT_NEGATIVE)
    CF_WARNING_FILL  = PatternFill(bgColor=palette.get("CF_WARNING_BG", "FEF9E7"))
    CF_WARNING_FONT  = Font(color=ACCENT_WARNING)

    _ACTIVE_STYLE = style


# ============================================================
# §3  Column Width Map
# ============================================================

COLUMN_WIDTHS = {
    "margin":      3,     # A col whitespace
    "id_short":    8,     # #, ID
    "name_cn":    16,     # Chinese name (2-4 chars)
    "name_en":    22,     # English name
    "description": 32,    # long text
    "number":     14,     # currency, amount
    "percentage": 12,     # %
    "date":       14,     # YYYY-MM-DD
    "status":     12,     # short label
}


# ============================================================
# §4  Number Formats
# ============================================================

FORMATS = {
    "integer":      "#,##0",
    "decimal_1":    "#,##0.0",
    "decimal_2":    "#,##0.00",
    "percentage":   "0.0%",
    "currency_cny": "¥#,##0.00",
    "currency_usd": "$#,##0.00",
    "date":         "YYYY-MM-DD",
}


# ============================================================
# §5  Style Factories
# ============================================================

def font_title():
    """16pt title font — left-aligned, no fill."""
    return Font(name=FONT_NAME, size=16, bold=HEADER_BOLD, color=PRIMARY)

def font_header():
    """11pt header font — text color on primary background."""
    return Font(name=FONT_NAME, size=11, bold=HEADER_BOLD, color=HEADER_TEXT)

def font_subheader():
    """11pt sub-header — primary color text."""
    return Font(name=FONT_NAME, size=11, bold=HEADER_BOLD, color=PRIMARY)

def font_body():
    """11pt body text."""
    return Font(name=FONT_NAME, size=11, color=NEUTRAL_900)

def font_caption():
    """9pt caption / footnote."""
    return Font(name=FONT_NAME, size=9, color=NEUTRAL_600)

def font_kpi():
    """22pt big KPI number."""
    return Font(name=FONT_NAME, size=22, bold=HEADER_BOLD, color=PRIMARY)

def font_kpi_label():
    """9pt KPI label."""
    return Font(name=FONT_NAME, size=9, color=NEUTRAL_600)


def make_chart_title(text, size_pt=12, bold=True, axis=False, max_line_chars=6):
    """
    Build a chart Title with font baked into <tx><rich><defRPr>/<rPr>.
    Ensures WPS and Office render identical font name and size.
    Uses FONT_NAME and HEADER_BOLD from §1 — no hardcoded font names.

    Args:
        axis: If True, set bodyPr rot=-5400000 (rotate -90°) for Y-axis titles.
        max_line_chars: For axis titles, auto-insert line breaks (\n) when text
              exceeds this length. Breaks at parentheses boundaries.
              The text stays in ONE run inside ONE paragraph — this prevents
              WPS/Office from creating separate overlapping text boxes.
              Set to 0 or None to disable.

    Key insight: Multiple <p> paragraphs in axis titles cause WPS to render
    them as stacked overlapping text boxes. Instead, we use a SINGLE <r> run
    with \\n line breaks inside the text, which both Office and WPS render
    as line breaks within the same text box.
    """
    from openpyxl.chart.title import Title
    from openpyxl.chart.text import Text, RichText
    from openpyxl.drawing.text import (
        Paragraph, ParagraphProperties, CharacterProperties,
        Font as DrawingFont, RichTextProperties, RegularTextRun,
        LineBreak,
    )
    from copy import deepcopy
    import re

    rpr = CharacterProperties(
        latin=DrawingFont(typeface=FONT_NAME),
        ea=DrawingFont(typeface=FONT_NAME),
        sz=int(size_pt * 100),
        b=bold if HEADER_BOLD else False,
    )

    def _insert_breaks(text, max_chars):
        """Insert \\n before parentheses when text exceeds max_chars."""
        if not max_chars or len(text) <= max_chars:
            return text
        # Insert \n before '(' or '（'
        result = re.sub(r'(?=[（(])', '\n', text, count=1)
        return result

    # For axis titles, insert line breaks to prevent overlap
    display_text = text
    if axis and max_line_chars:
        display_text = _insert_breaks(text, max_line_chars)

    # Single paragraph, single run with \n inside the text.
    # Both Office and WPS render \n as line breaks within one text box.
    # Do NOT use multiple <p> paragraphs — WPS renders them as separate
    # overlapping text boxes on axis titles.
    run = RegularTextRun(rPr=deepcopy(rpr), t=display_text)

    inner_body = RichTextProperties(rot=-5400000) if axis else RichTextProperties()
    para = Paragraph(
        pPr=ParagraphProperties(defRPr=deepcopy(rpr)),
        r=[run],
    )
    rich = RichText(bodyPr=inner_body, p=[para])

    # Outer txPr: Office reads rotation from here for axis titles
    if axis:
        outer_body = RichTextProperties(rot=-5400000)
        txPr = RichText(
            bodyPr=outer_body,
            p=[Paragraph(pPr=ParagraphProperties(defRPr=deepcopy(rpr)))],
        )
        return Title(tx=Text(rich=rich), txPr=txPr)
    return Title(tx=Text(rich=rich))


def fill_header():
    return PatternFill("solid", fgColor=PRIMARY)

def fill_total():
    return PatternFill("solid", fgColor=SECONDARY)

def fill_data_row(row_index: int):
    """Alternating row: even=white, odd=warm-white."""
    color = NEUTRAL_0 if row_index % 2 == 0 else NEUTRAL_100
    return PatternFill("solid", fgColor=color)


def border_header():
    """Thin bottom border under header row."""
    return Border(bottom=Side(style="thin", color=NEUTRAL_200))

def border_total():
    """Medium top border above totals row."""
    return Border(top=Side(style="medium", color=NEUTRAL_200))


def align_title():
    return Alignment(horizontal="left", vertical="center")

def align_header():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def align_number():
    return Alignment(horizontal="right", vertical="center", wrap_text=True)

def align_text():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def align_date():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)


# ============================================================
# §6  Sheet Setup Helpers
# ============================================================

ROW_HEIGHTS = {
    "margin":   15,   # row 1 top whitespace
    "title":    32,   # row 2
    "spacer":    8,   # row 3
    "header":   28,   # row 4
    "data":     22,   # data rows
    "total":    26,   # totals row
}


def setup_sheet(ws, title: str = None, last_col: int = None):
    """
    Apply standard sheet setup:
      - hide grid lines
      - set margin column A width
      - set row 1/2/3 heights
      - optionally write & style title at B2
    """
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = COLUMN_WIDTHS["margin"]
    ws.row_dimensions[1].height = ROW_HEIGHTS["margin"]
    ws.row_dimensions[2].height = ROW_HEIGHTS["title"]
    ws.row_dimensions[3].height = ROW_HEIGHTS["spacer"]

    if title and last_col:
        ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last_col)
        cell = ws.cell(row=2, column=2, value=title)
        cell.font = font_title()
        cell.alignment = align_title()


def style_header_row(ws, row_num: int, col_start: int, col_end: int):
    """Apply header style to a row range."""
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = fill_header()
        cell.font = font_header()
        cell.alignment = align_header()
        cell.border = border_header()
    ws.row_dimensions[row_num].height = ROW_HEIGHTS["header"]


def style_data_row(ws, row_num: int, col_start: int, col_end: int, row_index: int):
    """Apply data row style (alternating fill). Enables wrap_text to prevent clipping."""
    fill = fill_data_row(row_index)
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = fill
        cell.font = font_body()
        # Preserve existing alignment but ensure wrap_text is enabled
        existing = cell.alignment
        if existing and (existing.horizontal or existing.vertical):
            cell.alignment = Alignment(
                horizontal=existing.horizontal or "left",
                vertical=existing.vertical or "center",
                wrap_text=True,
            )
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[row_num].height = ROW_HEIGHTS["data"]


def style_total_row(ws, row_num: int, col_start: int, col_end: int):
    """Apply totals row style."""
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = fill_total()
        cell.font = font_subheader()
        cell.border = border_total()
    ws.row_dimensions[row_num].height = ROW_HEIGHTS["total"]


# ============================================================
# §6.1  Chart Factory Functions
# ============================================================

def create_bar_chart(chart_type="col", grouping="clustered", gap_width=80,
                     overlap=100, style=10, width=18, height=10, **kwargs):
    """
    Create a BarChart with sane defaults that prevent the "thin bar" / offset bug.

    Key fixes baked in:
      - gapWidth=80 (default 150 → bars too thin)
      - overlap=100 (bars fill their slot, no empty gap for line series)

    Returns an openpyxl BarChart ready for add_data / set_categories.
    """
    from openpyxl.chart import BarChart
    chart = BarChart()
    chart.type = chart_type
    chart.grouping = grouping
    chart.gapWidth = gap_width
    chart.overlap = overlap
    chart.style = style
    chart.width = width
    chart.height = height
    return chart


def create_line_chart(style=10, width=18, height=11, **kwargs):
    """Create a LineChart with standard defaults."""
    from openpyxl.chart import LineChart
    chart = LineChart()
    chart.style = style
    chart.width = width
    chart.height = height
    return chart


def create_pie_chart(style=10, width=14, height=10, **kwargs):
    """Create a PieChart with standard defaults."""
    from openpyxl.chart import PieChart
    chart = PieChart()
    chart.style = style
    chart.width = width
    chart.height = height
    return chart


def setup_chart_titles(chart, title=None, y_title=None, x_title=None,
                       title_size=12, axis_size=10):
    """
    Set chart title and axis titles using make_chart_title() for
    cross-platform font consistency (Office + WPS).

    This is the ONLY correct way to set chart titles. Never do:
        chart.title = "some string"        # ← WRONG
        chart.y_axis.title = "some string" # ← WRONG

    Args:
        chart: openpyxl chart object
        title: main chart title (optional)
        y_title: Y-axis title (optional, auto-rotated -90°)
        x_title: X-axis title (optional)
        title_size: font size for main title (default 12)
        axis_size: font size for axis titles (default 10)
    """
    if title is not None:
        chart.title = make_chart_title(title, size_pt=title_size, bold=True)
    if y_title is not None:
        chart.y_axis.title = make_chart_title(y_title, size_pt=axis_size, bold=False, axis=True)
    if x_title is not None:
        chart.x_axis.title = make_chart_title(x_title, size_pt=axis_size, bold=False)


def apply_chart_colors(chart, colors=None):
    """
    Apply palette colors to all series in a chart.
    Call AFTER add_data().

    Args:
        chart: openpyxl chart object (BarChart, LineChart, etc.)
        colors: list of hex color strings (default: CHART_COLORS)
    """
    if colors is None:
        colors = CHART_COLORS
    for i, series in enumerate(chart.series):
        color_hex = colors[i % len(colors)]
        series.graphicalProperties.solidFill = color_hex
        # For line charts, also set line color
        if hasattr(series.graphicalProperties, 'line') and series.graphicalProperties.line is not None:
            series.graphicalProperties.line.solidFill = color_hex


def apply_pie_colors(chart, count, colors=None):
    """
    Apply palette colors to pie chart data points.
    Call AFTER add_data().

    Args:
        chart: openpyxl PieChart
        count: number of data points (slices)
        colors: list of hex color strings (default: CHART_COLORS)
    """
    from openpyxl.chart.series import DataPoint
    if colors is None:
        colors = CHART_COLORS
    for idx in range(count):
        pt = DataPoint(idx=idx)
        pt.graphicalProperties.solidFill = colors[idx % len(colors)]
        chart.series[0].data_points.append(pt)


# ============================================================
# §7  Utility Functions
# ============================================================

def normalize_cell_value(value):
    """Normalize cell values: convert invisible whitespace variants to None."""
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip().replace("\xa0", "").replace("\u200b", "")
        if stripped == "":
            return None
    return value


def copy_style(source_cell, target_cell):
    """Copy all styling from source to target cell."""
    target_cell.font = copy(source_cell.font)
    target_cell.fill = copy(source_cell.fill)
    target_cell.border = copy(source_cell.border)
    target_cell.alignment = copy(source_cell.alignment)
    target_cell.number_format = source_cell.number_format


def auto_fit_row_heights(ws, header_row=None, data_start_row=None, data_end_row=None,
                         base_height=22, line_height=16, font_size=11, min_height=22):
    """
    Auto-fit row heights based on cell content and column width.
    Prevents text from being clipped when wrap_text causes multi-line display.

    Call AFTER auto_fit_columns() (needs final column widths to calculate wrapping).

    Args:
        ws: worksheet
        header_row: header row number (auto-detected if None)
        data_start_row: first data row (auto-detected as header_row + 1 if None)
        data_end_row: last data row (auto-detected as ws.max_row if None)
        base_height: base row height for single-line content (default 22)
        line_height: additional height per extra line (default 16)
        font_size: font size in pt for estimation (default 11)
        min_height: minimum row height (default 22)
    """
    import unicodedata
    import math

    def _display_width(text):
        """Estimate display width: CJK chars count as ~1.7, others as 1."""
        if text is None:
            return 0
        s = str(text)
        w = 0
        for ch in s:
            if unicodedata.east_asian_width(ch) in ('W', 'F'):
                w += 1.7
            else:
                w += 1
        return w

    def _estimate_lines(text, col_width):
        """Estimate how many lines a cell value needs given column width."""
        if text is None:
            return 1
        s = str(text)
        # Count explicit line breaks
        lines = s.split('\n')
        total_lines = 0
        # Available char width ≈ col_width * 0.9 (accounting for cell padding)
        avail_chars = max(1, col_width * 0.9)
        for line in lines:
            line_w = _display_width(line)
            if line_w <= avail_chars:
                total_lines += 1
            else:
                total_lines += math.ceil(line_w / avail_chars)
        return max(1, total_lines)

    # Auto-detect header row
    if header_row is None:
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row=row, column=2).value
            if val is not None:
                header_row = row
                break
        if header_row is None:
            return

    if data_start_row is None:
        data_start_row = header_row + 1
    if data_end_row is None:
        data_end_row = ws.max_row

    # Get column widths
    col_widths = {}
    for col in range(1, ws.max_column + 1):
        col_letter = ws.cell(row=1, column=col).column_letter
        dim = ws.column_dimensions.get(col_letter)
        if dim and dim.width:
            col_widths[col] = dim.width
        else:
            col_widths[col] = 11  # openpyxl default

    # Adjust header row height
    max_header_lines = 1
    for col in range(2, ws.max_column + 1):
        cell = ws.cell(row=header_row, column=col)
        if cell.value and (cell.alignment and cell.alignment.wrap_text):
            lines = _estimate_lines(cell.value, col_widths.get(col, 11))
            max_header_lines = max(max_header_lines, lines)
    if max_header_lines > 1:
        header_height = base_height + (max_header_lines - 1) * line_height
        ws.row_dimensions[header_row].height = max(ROW_HEIGHTS["header"], header_height)

    # Adjust data rows
    for row_num in range(data_start_row, data_end_row + 1):
        max_lines = 1
        for col in range(2, ws.max_column + 1):
            cell = ws.cell(row=row_num, column=col)
            if cell.value is None:
                continue
            # Check if cell has wrap_text enabled
            has_wrap = (cell.alignment and cell.alignment.wrap_text)
            # Also check if content has explicit newlines
            has_newlines = isinstance(cell.value, str) and '\n' in cell.value
            if has_wrap or has_newlines:
                lines = _estimate_lines(cell.value, col_widths.get(col, 11))
                max_lines = max(max_lines, lines)
        if max_lines > 1:
            needed_height = base_height + (max_lines - 1) * line_height
            ws.row_dimensions[row_num].height = max(min_height, needed_height)
        else:
            # Keep at least min_height (don't shrink below base)
            current = ws.row_dimensions[row_num].height
            if current is None or current < min_height:
                ws.row_dimensions[row_num].height = min_height


def auto_fit_columns(ws, min_width=8, max_width=28, header_row=None, data_start_row=None):
    """
    Auto-fit column widths based on DATA content (not header).
    Headers that exceed the computed width get wrap_text=True instead of stretching the column.

    Args:
        ws: worksheet
        min_width: minimum column width (default 8)
        max_width: maximum column width (default 28)
        header_row: row number of the header (auto-detected if None)
        data_start_row: first data row (auto-detected as header_row + 1 if None)
    """
    import unicodedata

    def _display_width(text):
        """Estimate display width: CJK chars count as ~1.7, others as 1."""
        if text is None:
            return 0
        s = str(text)
        w = 0
        for ch in s:
            if unicodedata.east_asian_width(ch) in ('W', 'F'):
                w += 1.7
            else:
                w += 1
        return w

    # Auto-detect header row: first row with data starting from column B
    if header_row is None:
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row=row, column=2).value
            if val is not None:
                header_row = row
                break
        if header_row is None:
            return

    if data_start_row is None:
        data_start_row = header_row + 1

    for col_cells in ws.iter_cols(min_col=1, max_col=ws.max_column, min_row=data_start_row, max_row=ws.max_row):
        if not col_cells:
            continue
        col_letter = col_cells[0].column_letter

        # Skip margin column A
        if col_letter == 'A':
            continue

        # Width based on data content only
        max_data_w = max((_display_width(c.value) for c in col_cells), default=0)
        width = min(max_width, max(min_width, max_data_w + 2))
        ws.column_dimensions[col_letter].width = width

        # If header text is wider than computed column width, wrap it
        header_cell = ws.cell(row=header_row, column=col_cells[0].column)
        header_w = _display_width(header_cell.value)
        if header_w > width:
            current_align = header_cell.alignment
            header_cell.alignment = Alignment(
                horizontal=current_align.horizontal or "center",
                vertical=current_align.vertical or "center",
                wrap_text=True,
            )

B = sys.modules[__name__]   # 单文件版：base 的定义已并入本模块命名空间

# 本机（Windows）存在微软雅黑，粗细适中可加粗；覆盖 base 的 Noto 回退
B.FONT_NAME = "Microsoft YaHei"
B.HEADER_BOLD = True

# ── 蓝色主题（用户指定：整体色调统一为蓝） ─────────────────────
# 强调色不再用绿/琥珀，改为单色相蓝阶：深 → 中 → 浅，
# Perm/PC/MS 在 KPI 卡、堆积柱状图、饼图中共享同一套蓝色身份色。
B.PRIMARY = "1B2A4A"          # 深藏青：表头、标题、主 KPI
B.PRIMARY_LIGHT = "D6E4F0"    # 浅蓝：筛选框、合计行、期间高亮
B.SECONDARY = "D6E4F0"
B.ACCENT_POSITIVE = "2E6DA4"  # 中蓝（原绿色位 → PC）
B.ACCENT_WARNING = "5B9BD5"   # 浅蓝（原琥珀位 → MS）
B.ACCENT_NEGATIVE = "8EAADB"  # 更浅蓝（备用）
B.NEUTRAL_900 = "1F2A36"      # 冷调正文
B.NEUTRAL_600 = "6E7B8B"      # 冷调次级文字
B.NEUTRAL_200 = "D9E2EC"      # 冷调分隔线
B.NEUTRAL_100 = "F2F6FA"      # 冷调斑马纹 / KPI 卡底
B.CHART_COLORS = ["1B2A4A", "2E6DA4", "5B9BD5", "8EAADB", "6E7B8B"]

# ============ 常量 ============
S_DASH, S_GRID, S_DP, S_DL, S_DOC, S_CALC = (
    "Dashboard", "Monthly Detail", "Data_Profile", "Data_Leavers", "Guide", "Calc")

QDP, QDL, QDOC, QC = f"'{S_DP}'", f"'{S_DL}'", f"'{S_DOC}'", f"'{S_CALC}'"

DATA_CAP = 5000            # 每张数据表可容纳的数据行数
R0, R1 = 5, 5 + DATA_CAP - 1   # 数据起始行 / 结束行 (5..5004)
HDR_ROW = 4                # 数据表表头所在行
N_MONTHS = 48              # 月度明细覆盖月数

PROFILE_HEADERS = [
    "StaffID", "1BankID", "Status", "Full Name", "Chinese Name",
    "Platform working on (currently", "Squad working on (currently",
    "Vendor", "Type", "Level", "JF with function", "Job function",
    "Gender", "Perm / PC / MS", "Role", "Tenure",
    "Staff start date", "Contract end date", "IBG/FRT", "Reporting manager.",
    "PLATFORM (Billing)", "PC code", "Dept CODE", "Last workday.",
]
# 关键列在一维列表中的下标（0 基），用于写测试数据
COL_ID, COL_STATUS = 0, 2
COL_TYPE = 13            # 招聘分类列 = 'Perm / PC / MS'（不是 'Type' 列）
COL_START, COL_CED, COL_LWD = 16, 17, 23

CATS = ["Perm", "PC", "MS"]

# 关键单元格
CELL_FILT = "C5"           # 看板类型筛选下拉框
CELL_PERIOD = "E5"         # 看板时间筛选下拉框（All / 年份 / 月份）
GRID_FIRST, GRID_LAST = 5, 5 + N_MONTHS - 1     # 月度明细数据行 5..52
GRID_TOTAL = GRID_LAST + 1


def q(name):
    return f"'{name}'"


def add_name(wb, name, ref):
    dn = DefinedName(name, attr_text=ref)
    try:
        wb.defined_names[name] = dn          # openpyxl >= 3.1
    except TypeError:
        wb.defined_names.append(dn)          # 兼容旧版


# ============ 名称定义 ============
def define_names(wb):
    def idx(sheet_q, header_cfg):
        return (f"INDEX({sheet_q}!$A${R0}:$AZ${R1},0,"
                f"MATCH({header_cfg},{sheet_q}!$4:$4,0))")

    add_name(wb, "cfgStartHdr", f"{QDOC}!$C$20")
    add_name(wb, "cfgLWDHdr",   f"{QDOC}!$C$21")
    add_name(wb, "cfgTypeHdr",  f"{QDOC}!$C$22")
    add_name(wb, "cfgStatusHdr", f"{QDOC}!$C$23")
    add_name(wb, "cfgIDHdr",    f"{QDOC}!$C$24")
    add_name(wb, "FiltCat",     f"'{S_DASH}'!${CELL_FILT[0]}${CELL_FILT[1:]}")
    add_name(wb, "FiltPeriod",  f"'{S_DASH}'!${CELL_PERIOD[0]}${CELL_PERIOD[1:]}")

    add_name(wb, "P_ID",    idx(QDP, "cfgIDHdr"))
    add_name(wb, "P_Start", idx(QDP, "cfgStartHdr"))
    add_name(wb, "P_LWD",   idx(QDP, "cfgLWDHdr"))
    add_name(wb, "P_Type",  idx(QDP, "cfgTypeHdr"))
    add_name(wb, "P_Status", idx(QDP, "cfgStatusHdr"))
    add_name(wb, "L_ID",    idx(QDL, "cfgIDHdr"))
    add_name(wb, "L_Start", idx(QDL, "cfgStartHdr"))
    add_name(wb, "L_LWD",   idx(QDL, "cfgLWDHdr"))
    add_name(wb, "L_Type",  idx(QDL, "cfgTypeHdr"))
    add_name(wb, "L_Status", idx(QDL, "cfgStatusHdr"))
    add_name(wb, "L_InP",   f"{QC}!$A${R0}:$A${R1}")
    add_name(wb, "P_TypeN", f"{QC}!$B${R0}:$B${R1}")
    add_name(wb, "L_TypeN", f"{QC}!$D${R0}:$D${R1}")
    add_name(wb, "P_StatusN", f"{QC}!$E${R0}:$E${R1}")
    add_name(wb, "L_StatusN", f"{QC}!$G${R0}:$G${R1}")
    add_name(wb, "L_TwinIna", f"{QC}!$J${R0}:$J${R1}")
    add_name(wb, "PeriodList", f"{QC}!$F$5:$F$58")
    add_name(wb, "SelStart", f"{QC}!$K$2")
    add_name(wb, "SelEnd",   f"{QC}!$K$3")
    add_name(wb, "SelStartTxt", f'IF({QC}!$K$2="","",TEXT({QC}!$K$2,"yyyy-mm"))')
    add_name(wb, "SelEndTxt",   f'IF({QC}!$K$3="","",TEXT({QC}!$K$3,"yyyy-mm"))')


# ============ 公式片段 ============
HFLAG = f"{QC}!$H$2+{QC}!$I$2>0"
NO_DATA = f"{QC}!$C$5=\"\""

PFILT = '(P_TypeN=UPPER(FiltCat))+(FiltCat="All")'
LFILT = '(L_TypeN=UPPER(FiltCat))+(FiltCat="All")'


def cat_expr_p(literal=None):
    return PFILT if literal is None else f'(P_TypeN="{literal.upper()}")'


def cat_expr_l(literal=None):
    return LFILT if literal is None else f'(L_TypeN="{literal.upper()}")'


def f_headcount(m, pexpr, lexpr):
    """期末在职：Active 一直算在职；Inactive 算到 Last Work Day 当月末（当日含）。"""
    eom = f"EOMONTH({m},0)"
    p = (f"SUMPRODUCT((ISNUMBER(P_Start))*(P_Start<={eom})"
         f"*(((P_StatusN=\"ACTIVE\")+((ISNUMBER(P_LWD))*(P_LWD>={eom})))>0)"
         f"*({pexpr}))")
    l_mask = (f"(ISNUMBER(L_Start))*(L_Start<={eom})"
              f"*(((L_StatusN=\"ACTIVE\")+((ISNUMBER(L_LWD))*(L_LWD>={eom})))>0)"
              f"*({lexpr})")
    return f"{p}+SUMPRODUCT({l_mask})-SUMPRODUCT({l_mask}*(L_InP=1))"


def f_outflow(m, pexpr, lexpr):
    """当月流出：Status=Inactive 且 Last Work Day 落在该月。
    两表重复时：只有 Profile 副本也是 Inactive 才去重（离职以 Leavers 表为准）。"""
    nxt = f"EDATE({m},1)"
    p = (f"SUMPRODUCT((P_StatusN=\"INACTIVE\")*(ISNUMBER(P_LWD))"
         f"*(P_LWD>={m})*(P_LWD<{nxt})*({pexpr}))")
    l_mask = (f"(L_StatusN=\"INACTIVE\")*(ISNUMBER(L_LWD))"
              f"*(L_LWD>={m})*(L_LWD<{nxt})*({lexpr})")
    return (f"{p}+SUMPRODUCT({l_mask})"
            f"-SUMPRODUCT({l_mask}*(L_InP=1)*(L_TwinIna=1))")


def f_join(m, pexpr, lexpr):
    """当月入职：Start Date 落在该月，两表合并去重"""
    nxt = f"EDATE({m},1)"
    p = f"SUMPRODUCT((ISNUMBER(P_Start))*(P_Start>={m})*(P_Start<{nxt})*({pexpr}))"
    l_mask = f"(ISNUMBER(L_Start))*(L_Start>={m})*(L_Start<{nxt})*({lexpr})"
    return f"{p}+SUMPRODUCT({l_mask})-SUMPRODUCT({l_mask}*(L_InP=1))"


def f_now_status(pexpr, lexpr):
    """当前在职：Status=Active 且入职日期已到（≤ 今天）。"""
    p = (f"SUMPRODUCT((P_StatusN=\"ACTIVE\")*(ISNUMBER(P_Start))"
         f"*(P_Start<=TODAY())*({pexpr}))")
    l_mask = (f"(L_StatusN=\"ACTIVE\")*(ISNUMBER(L_Start))"
              f"*(L_Start<=TODAY())*({lexpr})")
    return f"{p}+SUMPRODUCT({l_mask})-SUMPRODUCT({l_mask}*(L_InP=1))"


def f_leavers_all(pexpr, lexpr):
    """全部时间流出 = Status=Inactive 的人数（Last Work Day 缺失也算）。"""
    p = f"SUMPRODUCT((P_StatusN=\"INACTIVE\")*({pexpr}))"
    l_mask = f"(L_StatusN=\"INACTIVE\")*({lexpr})"
    return (f"{p}+SUMPRODUCT({l_mask})"
            f"-SUMPRODUCT({l_mask}*(L_InP=1)*(L_TwinIna=1))")


def f_leavers_period(pexpr, lexpr):
    """期间流出：Status=Inactive 且 Last Work Day 落在 [SelStart, SelEnd]。"""
    p = (f"SUMPRODUCT((P_StatusN=\"INACTIVE\")*(ISNUMBER(P_LWD))"
         f"*(P_LWD>=SelStart)*(P_LWD<=SelEnd)*({pexpr}))")
    l_mask = (f"(L_StatusN=\"INACTIVE\")*(ISNUMBER(L_LWD))"
              f"*(L_LWD>=SelStart)*(L_LWD<=SelEnd)*({lexpr})")
    return (f"{p}+SUMPRODUCT({l_mask})"
            f"-SUMPRODUCT({l_mask}*(L_InP=1)*(L_TwinIna=1))")


def f_join_period(pexpr, lexpr):
    """期间入职：Start Date 落在 [SelStart, SelEnd]，两表合并去重"""
    p = (f"SUMPRODUCT((ISNUMBER(P_Start))"
         f"*(P_Start>=SelStart)*(P_Start<=SelEnd)*({pexpr}))")
    l_mask = (f"(ISNUMBER(L_Start))"
              f"*(L_Start>=SelStart)*(L_Start<=SelEnd)*({lexpr})")
    return f"{p}+SUMPRODUCT({l_mask})-SUMPRODUCT({l_mask}*(L_InP=1))"


# ============ 通用样式小工具 ============
def caption_cell(ws, ref, text, merge_to=None):
    c = ws[ref]
    c.value = text
    c.font = B.font_caption()
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    if merge_to:
        ws.merge_cells(f"{ref}:{merge_to}")
    return c


def subheader_cell(ws, ref, text):
    c = ws[ref]
    c.value = text
    c.font = Font(name=B.FONT_NAME, size=12, bold=B.HEADER_BOLD, color=B.PRIMARY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    return c


def style_axis(axis, pos, num_fmt=None):
    """显式写出轴属性；OOXML 布尔缺省按 true 处理，不写 delete=False 轴会被隐藏。"""
    axis.delete = False
    axis.axPos = pos
    axis.tickLblPos = "nextTo"
    axis.majorTickMark = "out"
    axis.minorTickMark = "none"
    if num_fmt:
        axis.number_format = num_fmt
    axis.txPr = RichText(
        bodyPr=RichTextProperties(),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(
            sz=900, solidFill=B.NEUTRAL_900,
            latin=DrawingFont(typeface=B.FONT_NAME),
            ea=DrawingFont(typeface=B.FONT_NAME))))])


def set_cat_labels(chart, sheet, col, r1, r2):
    """文本类目必须用 strRef，否则 Excel 显示空白类目。"""
    from openpyxl.chart.data_source import AxDataSource, StrRef
    ref = f"'{sheet.title}'!${col}${r1}:${col}${r2}"
    for s in chart.series:
        s.cat = AxDataSource(strRef=StrRef(f=ref))


# ============ 工作簿骨架 ============
def build_workbook():
    wb = Workbook()
    dash = wb.active
    dash.title = S_DASH
    grid = wb.create_sheet(S_GRID)
    dp = wb.create_sheet(S_DP)
    dl = wb.create_sheet(S_DL)
    doc = wb.create_sheet(S_DOC)
    calc = wb.create_sheet(S_CALC)

    dash.sheet_properties.tabColor = B.PRIMARY
    grid.sheet_properties.tabColor = "8496B0"
    dp.sheet_properties.tabColor = "8496B0"
    dl.sheet_properties.tabColor = "8496B0"
    doc.sheet_properties.tabColor = "2E6DA4"
    calc.sheet_state = "hidden"

    build_data_sheet(dp, "Data_Profile  (paste '1. IBGT Profile' rows here)",
                     "Copy the data from your '1. IBGT Profile' sheet and paste it here - headers at B4, "
                     "or data only at B5. Prefer Paste Values so dates stay real dates.")
    build_data_sheet(dl, "Data_Leavers  (paste Leaver history here)",
                     "Copy all historical leavers from your '1.1 Leavers' sheet - headers at B4, or data only at B5. "
                     "This sheet may stay empty if all leavers are already included in Data_Profile.")
    build_grid_sheet(grid)
    build_dash_sheet(dash)
    build_doc_sheet(doc)
    build_calc_sheet(calc)
    define_names(wb)

    wb.calculation.fullCalcOnLoad = True
    wb.properties.creator = "Z.ai"
    return wb


# ============ 数据粘贴表 ============
def build_data_sheet(ws, title, note):
    last_col = 1 + len(PROFILE_HEADERS)          # B..P → 16
    B.setup_sheet(ws, title=title, last_col=last_col)
    caption_cell(ws, "B3", note, merge_to=f"{get_column_letter(last_col)}3")
    ws.row_dimensions[3].height = 26

    for i, h in enumerate(PROFILE_HEADERS):
        ws.cell(row=HDR_ROW, column=2 + i, value=h)
    B.style_header_row(ws, HDR_ROW, 2, last_col)

    # 斑马纹模板区（前 500 行），日期列格式覆盖到 5000 行
    date_cols = {2 + COL_START, 2 + COL_CED, 2 + COL_LWD}
    type_col = 2 + COL_TYPE
    zebra_rows = 500
    for r in range(R0, R0 + zebra_rows):
        B.style_data_row(ws, r, 2, last_col, r - R0)
    for c in sorted(date_cols):
        for r in range(R0, R1 + 1):
            cell = ws.cell(row=r, column=c)
            cell.number_format = "yyyy-mm-dd"
            cell.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(R0, R1 + 1):
        ws.cell(row=r, column=type_col).alignment = Alignment(
            horizontal="center", vertical="center")

    widths = [10, 10, 10, 20, 13, 24, 22, 14, 10, 9, 15, 17, 10, 14, 12, 8, 13, 13, 9, 18, 16, 9, 10, 13]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(2 + i)].width = w
    ws.freeze_panes = "B5"

    # 打印版式：只打印首页（粘贴模板区不需要整表打印）
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = "A1:Y54"


# ============ 月度明细 ============
def build_grid_sheet(ws):
    B.setup_sheet(ws, title="Monthly Turnover Detail  (linked to the Dashboard filter)", last_col=11)
    caption_cell(
        ws, "B3",
        "Rules: Headcount EOM = employed at month end (the Last Work Day itself counts as an active day); "
        "Leavers = Status Inactive with Last Work Day in that month; Joiners = start date in that month. "
        "C/D/H follow the Dashboard filter; Perm/PC/MS columns always show the full breakdown; "
        "48 months from the earliest start date.",
        merge_to="K3")
    ws.row_dimensions[3].height = 58

    headers = ["Month", "Headcount EOM\n(filtered)", "Leavers\n(filtered)",
               "Perm Leavers", "PC Leavers", "MS Leavers", "Joiners\n(filtered)"]
    for i, h in enumerate(headers):
        ws.cell(row=4, column=2 + i, value=h)
    B.style_header_row(ws, 4, 2, 8)

    flags = f"{QC}!$H$2+{QC}!$I$2>0"
    for r in range(GRID_FIRST, GRID_LAST + 1):
        m = f"{QC}!$C{r}"
        guard = f'IF({flags},"",IF({m}="","",'
        row = r - GRID_FIRST
        ws.cell(row=r, column=2, value=f'={guard}TEXT({m},"yyyy-mm")))')
        ws.cell(row=r, column=3, value=f"={guard}{f_headcount(m, PFILT, LFILT)}))")
        ws.cell(row=r, column=4, value=f"={guard}{f_outflow(m, PFILT, LFILT)}))")
        ws.cell(row=r, column=5, value=f"={guard}{f_outflow(m, cat_expr_p('Perm'), cat_expr_l('Perm'))}))")
        ws.cell(row=r, column=6, value=f"={guard}{f_outflow(m, cat_expr_p('PC'), cat_expr_l('PC'))}))")
        ws.cell(row=r, column=7, value=f"={guard}{f_outflow(m, cat_expr_p('MS'), cat_expr_l('MS'))}))")
        ws.cell(row=r, column=8, value=f"={guard}{f_join(m, PFILT, LFILT)}))")
        B.style_data_row(ws, r, 2, 8, row)
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(3, 9):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "#,##0"

    # 合计行
    tr = GRID_TOTAL
    ws.cell(row=tr, column=2, value="Total")
    ws.cell(row=tr, column=3, value="—")
    for c in range(4, 9):
        letter = get_column_letter(c)
        ws.cell(row=tr, column=c, value=f"=SUM({letter}{GRID_FIRST}:{letter}{GRID_LAST})")
    B.style_total_row(ws, tr, 2, 8)
    for c in range(4, 9):
        cell = ws.cell(row=tr, column=c)
        cell.number_format = "#,##0"
        cell.alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=tr, column=2).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=tr, column=3).alignment = Alignment(horizontal="center", vertical="center")

    # 右侧：当前在职构成（饼图数据源）
    ws.cell(row=4, column=10, value="Type")
    ws.cell(row=4, column=11, value="Current Headcount")
    B.style_header_row(ws, 4, 10, 11)
    for i, cat in enumerate(CATS):
        r = 5 + i
        ws.cell(row=r, column=10, value=cat)
        ws.cell(row=r, column=11, value=f"='{S_DASH}'!{'BEH'[i]}$11")
        B.style_data_row(ws, r, 10, 11, i)
        ws.cell(row=r, column=10).alignment = Alignment(horizontal="center", vertical="center")
        k = ws.cell(row=r, column=11)
        k.alignment = Alignment(horizontal="right", vertical="center")
        k.number_format = "#,##0"
    ws.cell(row=8, column=10, value="Total")
    ws.cell(row=8, column=11, value="=SUM(K5:K7)")
    B.style_total_row(ws, 8, 10, 11)
    k = ws.cell(row=8, column=11)
    k.alignment = Alignment(horizontal="right", vertical="center")
    k.number_format = "#,##0"
    ws.cell(row=8, column=10).alignment = Alignment(horizontal="center", vertical="center")

    caption_cell(ws, "J10", "Referenced by the pie chart - do not delete", merge_to="K10")

    # 选中期间在明细表中高亮（All 时不高亮）
    ws.conditional_formatting.add(
        f"B{GRID_FIRST}:H{GRID_LAST}",
        FormulaRule(
            formula=[f'AND($B{GRID_FIRST}<>"",'
                     f'$B{GRID_FIRST}>=SelStartTxt,$B{GRID_FIRST}<=SelEndTxt)'],
            fill=PatternFill(start_color=B.SECONDARY, end_color=B.SECONDARY,
                             fill_type="solid")))

    ws.column_dimensions["B"].width = 10
    for c in "CDEFGH":
        ws.column_dimensions[c].width = 12
    ws.column_dimensions["I"].width = 3
    ws.column_dimensions["J"].width = 10
    ws.column_dimensions["K"].width = 13
    ws.row_dimensions[4].height = 40
    ws.freeze_panes = "C5"

    # 打印/导出版式：横向一页，覆盖全部 48 个月 + 合计行
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:L{GRID_TOTAL + 1}"
    ws.print_title_rows = "4:4"


# ============ 看板 ============
def build_dash_sheet(ws):
    B.setup_sheet(ws, title="IBGT Department Turnover Dashboard", last_col=17)
    for c in range(2, 18):
        ws.column_dimensions[get_column_letter(c)].width = 12

    # 动态数据源说明
    caption = ws["B3"]
    caption.value = (
        f'=IF({QC}!$H$2+{QC}!$I$2>0,"⚠ Header mapping broken - open the Guide sheet and check C20:C24",'
        f'"Data source: Data_Profile "&TEXT(SUMPRODUCT(--(P_ID<>"")),"0")&" rows  |  Data_Leavers "'
        f'&TEXT(SUMPRODUCT(--(L_ID<>"")),"0")&" rows  |  "'
        f'&IF({QC}!$C$5="","No data yet - paste data and it updates automatically",'
        f'TEXT({QC}!$C$5,"yyyy-mm")&" onward, 48 months"))')
    caption.font = B.font_caption()
    caption.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("B3:Q3")
    ws.row_dimensions[3].height = 16
    ws.row_dimensions[4].height = 6

    # 筛选行：类型 + 时间
    lab = ws["B5"]
    lab.value = "Type Filter"
    lab.font = B.font_subheader()
    lab.alignment = Alignment(horizontal="left", vertical="center")

    def filter_cell(ref, value, title, err):
        c = ws[ref]
        c.value = value
        c.font = Font(name=B.FONT_NAME, size=12, bold=True, color=B.PRIMARY)
        c.fill = PatternFill("solid", fgColor=B.SECONDARY)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(bottom=Side(style="thin", color=B.NEUTRAL_200))
        dv = DataValidation(type="list", formula1=title, allow_blank=False)
        dv.error = err
        dv.errorTitle = "Invalid selection"
        ws.add_data_validation(dv)
        dv.add(ref)

    filter_cell(CELL_FILT, "All", '"All,Perm,PC,MS"', "Please choose: All / Perm / PC / MS")
    filter_cell(CELL_PERIOD, "All", "PeriodList",
                "Please choose a period: All, a year, or a month (yyyy-mm)")

    plab = ws["D5"]
    plab.value = "Period"
    plab.font = B.font_subheader()
    plab.alignment = Alignment(horizontal="left", vertical="center")

    hint = ws["G5"]
    hint.value = ("Switch type and period - cards update instantly; headcount is measured "
                  "at the end of the selected period, joiners & leavers within it")
    hint.font = B.font_caption()
    hint.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("G5:Q5")
    ws.row_dimensions[5].height = 22
    ws.row_dimensions[6].height = 8

    # ---- KPI 卡片 ----
    def card(row_lab, col_start, label, formula, color=None, num_fmt="#,##0"):
        col_end = col_start + 1
        for r in (row_lab, row_lab + 1):
            for c in range(col_start, col_end + 1):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=B.NEUTRAL_100)
        lc = ws.cell(row=row_lab, column=col_start, value=label)
        lc.font = B.font_kpi_label()
        lc.alignment = Alignment(horizontal="left", vertical="bottom", indent=1)
        ws.merge_cells(start_row=row_lab, start_column=col_start,
                       end_row=row_lab, end_column=col_end)
        vc = ws.cell(row=row_lab + 1, column=col_start, value=formula)
        vc.font = Font(name=B.FONT_NAME, size=22, bold=True,
                       color=color or B.PRIMARY)
        vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        vc.number_format = num_fmt
        ws.merge_cells(start_row=row_lab + 1, start_column=col_start,
                       end_row=row_lab + 1, end_column=col_end)

    guard_kpi = f'IF({HFLAG},"Fix header mapping",IF({NO_DATA},"-",'
    sel_all = f'{QC}!$K$3=""'          # 未选具体期间 = All
    card(7, 2, "Period-End Headcount",
         f"={guard_kpi}IF({sel_all},{f_now_status(PFILT, LFILT)},"
         f"{f_headcount(f'{QC}!$K$3', PFILT, LFILT)})))")
    card(7, 5, "Joiners in Period",
         f"={guard_kpi}IF({sel_all},SUM({q(S_GRID)}!$H${GRID_FIRST}:$H${GRID_LAST}),"
         f"{f_join_period(PFILT, LFILT)})))")
    card(7, 8, "Leavers in Period",
         f"={guard_kpi}IF({sel_all},{f_leavers_all(PFILT, LFILT)},"
         f"{f_leavers_period(PFILT, LFILT)})))")
    card(7, 11, "Data Coverage",
         f'={guard_kpi}TEXT(MAX(P_Start,P_LWD,L_Start,L_LWD),"yyyy-mm")))')

    card(10, 2, "Headcount - Perm",
         f"={guard_kpi}IF({sel_all},{f_now_status(cat_expr_p('Perm'), cat_expr_l('Perm'))},"
         f"{f_headcount(f'{QC}!$K$3', cat_expr_p('Perm'), cat_expr_l('Perm'))})))",
         color=B.CHART_COLORS[0])
    card(10, 5, "Headcount - PC",
         f"={guard_kpi}IF({sel_all},{f_now_status(cat_expr_p('PC'), cat_expr_l('PC'))},"
         f"{f_headcount(f'{QC}!$K$3', cat_expr_p('PC'), cat_expr_l('PC'))})))",
         color=B.CHART_COLORS[1])
    card(10, 8, "Headcount - MS",
         f"={guard_kpi}IF({sel_all},{f_now_status(cat_expr_p('MS'), cat_expr_l('MS'))},"
         f"{f_headcount(f'{QC}!$K$3', cat_expr_p('MS'), cat_expr_l('MS'))})))",
         color=B.CHART_COLORS[2])
    latest_out = (f"INDEX({q(S_GRID)}!$D${GRID_FIRST}:$D${GRID_LAST},"
                  f"MATCH(DATE(YEAR(MAX(P_LWD,L_LWD)),MONTH(MAX(P_LWD,L_LWD)),1),"
                  f"{QC}!$C${GRID_FIRST}:$C${GRID_LAST},0))")
    card(10, 11, "Latest Month Leavers",
         f'={guard_kpi}IFERROR({latest_out},"-")))')
    for r, h in ((7, 16), (8, 34), (9, 10), (10, 16), (11, 34), (12, 8)):
        ws.row_dimensions[r].height = h

    # ---- 图表 ----
    grid = ws.parent[S_GRID]
    line = B.create_line_chart(width=16.8, height=9)
    B.setup_chart_titles(line, title="Month-End Headcount Trend (filtered)")
    line.add_data(Reference(grid, min_col=3, min_row=4, max_row=GRID_LAST),
                  titles_from_data=True)
    set_cat_labels(line, grid, "B", GRID_FIRST, GRID_LAST)
    s = line.series[0]
    s.smooth = False
    s.graphicalProperties.line.solidFill = B.PRIMARY
    s.graphicalProperties.line.width = 28575
    s.marker = Marker(symbol="circle", size=5)
    s.marker.graphicalProperties.solidFill = B.PRIMARY
    style_axis(line.x_axis, "b")
    style_axis(line.y_axis, "l", "#,##0")
    line.legend = None
    ws.add_chart(line, "B13")

    stack = B.create_bar_chart(grouping="stacked", overlap=100, width=16.8, height=9)
    B.setup_chart_titles(stack, title="Monthly Leavers by Type")
    stack.add_data(Reference(grid, min_col=5, max_col=7, min_row=4, max_row=GRID_LAST),
                   titles_from_data=True)
    set_cat_labels(stack, grid, "B", GRID_FIRST, GRID_LAST)
    for i, sr in enumerate(stack.series):
        sr.graphicalProperties.solidFill = B.CHART_COLORS[i]
    style_axis(stack.x_axis, "b")
    style_axis(stack.y_axis, "l", "#,##0")
    stack.y_axis.majorUnit = 1          # 人数为整数，避免 0.5 步长产生 0/0/1/1 刻度
    stack.legend.position = "b"
    stack.legend.overlay = False
    ws.add_chart(stack, "K13")

    pie = B.create_pie_chart(width=11.5, height=8.5)
    B.setup_chart_titles(pie, title="Current Headcount by Type")
    pie.add_data(Reference(grid, min_col=11, min_row=4, max_row=7),
                 titles_from_data=True)
    set_cat_labels(pie, grid, "J", 5, 7)
    B.apply_pie_colors(pie, 3)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showCatName = False
    pie.dataLabels.showSerName = False
    pie.dataLabels.showLegendKey = False
    pie.dataLabels.showBubbleSize = False
    pie.dataLabels.dLblPos = "bestFit"
    pie.legend.position = "r"
    pie.legend.overlay = False
    ws.add_chart(pie, "B33")

    # ---- 数据概览 ----
    subheader_cell(ws, "K33", "Data Overview")
    ov = [
        ('="Data_Profile rows: "&TEXT(SUMPRODUCT(--(P_ID<>"")),"0")&"    |    Data_Leavers rows: "'
         '&TEXT(SUMPRODUCT(--(L_ID<>"")),"0")'),
        ('="Profile rows with Last Work Day (leavers): "&TEXT(SUMPRODUCT((P_ID<>"")*ISNUMBER(P_LWD)),"0")'),
        ('="Leavers also present in Profile (auto-deduplicated): "&TEXT(SUMPRODUCT(--(L_InP=1)),"0")'),
        (f'=IFERROR("Date anomalies: Profile "&TEXT(SUMPRODUCT((P_ID<>"")*(P_Start<>"")*(ISNUMBER(P_Start)=FALSE)),"0")'
         f'&" start / "&TEXT(SUMPRODUCT((P_ID<>"")*(P_LWD<>"")*(ISNUMBER(P_LWD)=FALSE)),"0")&" LWD    |    "'
         f'&"Leavers "&TEXT(SUMPRODUCT((L_ID<>"")*(L_Start<>"")*(ISNUMBER(L_Start)=FALSE)),"0")'
         f'&" / "&TEXT(SUMPRODUCT((L_ID<>"")*(L_LWD<>"")*(ISNUMBER(L_LWD)=FALSE)),"0")&" LWD","Pending header mapping fix")'),
        ('="Inactive rows without Last Work Day (cannot be placed in a month): "'
         '&TEXT(SUMPRODUCT((P_StatusN="INACTIVE")*(P_LWD=""))'
         '+SUMPRODUCT((L_StatusN="INACTIVE")*(L_LWD="")),"0")'),
        '="Latest data month: "&IF(' + QC + '!$C$5="","-",'
        'TEXT(EOMONTH(MAX(P_Start,P_LWD,L_Start,L_LWD),0),"yyyy-mm"))',
    ]
    for i, f in enumerate(ov):
        r = 34 + i
        cell = ws.cell(row=r, column=11, value=f)
        cell.font = B.font_caption()
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=17)
        ws.row_dimensions[r].height = 16
    note = ws.cell(row=40, column=11,
                   value="Rules: an employee counts as active through their Last Work Day; Data_Leavers rows "
                         "without Last Work Day are excluded from headcount; Contract End Date is reserved. "
                         "See the Guide sheet.")
    note.font = B.font_caption()
    note.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.merge_cells("K40:Q41")
    note.fill = PatternFill("solid", fgColor=B.NEUTRAL_100)
    ws.row_dimensions[40].height = 16
    ws.row_dimensions[41].height = 16

    foot = ws.cell(row=52, column=2,
                   value="Paste data into Data_Profile / Data_Leavers and the dashboard recalculates automatically "
                         "(Ctrl+Alt+F9 forces a full recalc). Calculation rules and header mapping: see the Guide sheet.")
    foot.font = B.font_caption()
    foot.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("B52:Q52")

    # 打印/导出版式：横向一页
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = "A1:S53"


# ============ 说明 ============
def doc_line(ws, r, text, height=None):
    cell = ws.cell(row=r, column=2, value=text)
    cell.font = B.font_body()
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    if height:
        ws.row_dimensions[r].height = height
    return cell


def build_doc_sheet(ws):
    B.setup_sheet(ws, title="Guide - IBGT Turnover Dashboard", last_col=8)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = "A1:H40"
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 26
    for c in "EFGH":
        ws.column_dimensions[c].width = 14

    subheader_cell(ws, "B4", "Quick Start")
    steps = [
        "1. Open your 'IBGT Profile' workbook and copy the whole '1. IBGT Profile' table.",
        "2. Paste it into the Data_Profile sheet: with headers at B4, or data only at B5. "
        "Prefer Paste Values so dates stay real dates.",
        "3. Paste the Leavers history into Data_Leavers the same way "
        "(it may stay empty if all leavers are already included in Data_Profile).",
        "4. On the Dashboard, switch cell C5 between All / Perm / PC / MS - "
        "KPI cards and the monthly detail follow instantly.",
        "5. Everything recalculates automatically after data changes; press Ctrl+Alt+F9 to force a full recalc.",
    ]
    r = 5
    for s_ in steps:
        doc_line(ws, r, s_, height=30 if len(s_) > 70 else 20)
        r += 1

    subheader_cell(ws, "B11", "Calculation Rules")
    rules = [
        "- Current headcount: Status = Active and the start date has passed. Anyone with Status = Inactive "
        "counts as a leaver; the Contract End Date is not used.",
        "- Leaving month: taken from Last Work Day. Leavers (a month) = Status Inactive with Last Work Day "
        "in that month.",
        "- Headcount EOM (a month): start date <= month end, and the person is Active, or Inactive with "
        "Last Work Day >= month end (the Last Work Day itself still counts as an active day).",
        "- Joiners (a month): Staff start date falls within that month, merged and deduplicated the same way.",
        "- Period filter: All time, a whole year, or a single month. Headcount is measured at the end of the "
        "selected period (All = as of today); joiners & leavers are counted within the period. The selected "
        "rows are highlighted in Monthly Detail.",
        "- Type filter: matches the 'Perm / PC / MS' column, case-insensitive and trimmed.",
        "- If the same StaffID appears in both tables, it is counted once; for leavers the Data_Leavers "
        "record wins. Inactive rows without Last Work Day cannot be placed in a month and are flagged in "
        "Data Quality Checks.",
    ]
    r = 12
    for t in rules:
        doc_line(ws, r, t, height=30 if len(t) > 90 else 20)
        r += 1

    subheader_cell(ws, "B19", "Header Mapping (if your headers differ, edit column C; column D shows the match)")
    cfg_rows = [
        ("Start Date column header", "Staff start date", "cfgStartHdr"),
        ("Last Work Day column header", "Last workday.", "cfgLWDHdr"),
        ("Perm / PC / MS column header", "Perm / PC / MS", "cfgTypeHdr"),
        ("Status column header", "Status", "cfgStatusHdr"),
        ("Staff ID column header", "StaffID", "cfgIDHdr"),
    ]
    r = 20
    for label, default, _ in cfg_rows:
        ws.cell(row=r, column=2, value=label).font = B.font_body()
        c = ws.cell(row=r, column=3, value=default)
        c.font = Font(name=B.FONT_NAME, size=11, bold=True, color=B.PRIMARY)
        c.fill = PatternFill("solid", fgColor=B.NEUTRAL_100)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        if label != "Status column header":
            d = ws.cell(row=r, column=4, value=(
                f'=IFERROR("Profile col "&MATCH($C{r},{QDP}!$4:$4,0)&"   |   Leavers col "'
                f'&MATCH($C{r},{QDL}!$4:$4,0),"Not found - edit column C")'))
            d.font = B.font_caption()
            d.alignment = Alignment(horizontal="left", vertical="center")
        else:
            d = ws.cell(row=r, column=4, value="(not used by the current rules)")
            d.font = B.font_caption()
        ws.row_dimensions[r].height = 20
        r += 1

    subheader_cell(ws, "B26", "Data Quality Checks")
    checks = [
        ("Data_Profile rows", '=TEXT(SUMPRODUCT(--(P_ID<>"")),"0")&" rows"'),
        ("Data_Leavers rows", '=TEXT(SUMPRODUCT(--(L_ID<>"")),"0")&" rows"'),
        ("Profile rows with Last Work Day", '=TEXT(SUMPRODUCT((P_ID<>"")*ISNUMBER(P_LWD)),"0")&" leavers"'),
        ("Leavers also in Profile (deduped)", '=TEXT(SUMPRODUCT(--(L_InP=1)),"0")&" row(s)"'),
        ("Date anomalies (start / LWD)",
         f'=TEXT(SUMPRODUCT((P_ID<>"")*(P_Start<>"")*(ISNUMBER(P_Start)=FALSE)),"0")&" / "'
         f'&TEXT(SUMPRODUCT((P_ID<>"")*(P_LWD<>"")*(ISNUMBER(P_LWD)=FALSE)),"0")&" in Profile   |   "'
         f'&TEXT(SUMPRODUCT((L_ID<>"")*(L_Start<>"")*(ISNUMBER(L_Start)=FALSE)),"0")&" / "'
         f'&TEXT(SUMPRODUCT((L_ID<>"")*(L_LWD<>"")*(ISNUMBER(L_LWD)=FALSE)),"0")&" in Leavers"'),
        ("Rows with blank / other Status",
         '=IFERROR(TEXT(SUMPRODUCT((P_ID<>"")*(P_StatusN<>"ACTIVE")*(P_StatusN<>"INACTIVE"))'
         '+SUMPRODUCT((L_ID<>"")*(L_StatusN<>"ACTIVE")*(L_StatusN<>"INACTIVE")),"0")&" rows",'
         '"Pending header mapping fix")'),
    ]
    r = 27
    for label, f in checks:
        lc = ws.cell(row=r, column=2, value=label)
        lc.font = B.font_body()
        lc.alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=4, value=f)
        c.font = Font(name=B.FONT_NAME, size=11, bold=True, color=B.PRIMARY)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 20
        r += 1

    subheader_cell(ws, "B33", "FAQ")
    faqs = [
        "Q: Data changed?  A: Just paste over the data sheets; KPIs, tables and charts update automatically.",
        "Q: Need more than 48 months?  A: The dashboard shows 48 months from the earliest start date; "
        "ask to extend the range if needed.",
        "Q: Types other than Perm/PC/MS?  A: The Dashboard filter is fixed to All/Perm/PC/MS; "
        "align your data or ask to change the options.",
        "Q: What is the Calc sheet?  A: A hidden helper sheet (type normalization and cross-table dedup). "
        "Do not delete it; right-click a sheet tab > Unhide to view.",
    ]
    r = 34
    for t in faqs:
        doc_line(ws, r, t, height=30 if len(t) > 90 else 20)
        r += 1


# ============ 计算（隐藏辅助表） ============
def build_calc_sheet(ws):
    ws.cell(row=1, column=1, value="InProfile (auto)")
    ws.cell(row=1, column=2, value="Profile Type normalized (auto)")
    ws.cell(row=1, column=3, value="Month chain (auto)")
    ws.cell(row=1, column=4, value="Leavers Type normalized (auto)")
    ws.cell(row=1, column=5, value="Profile Status normalized (auto)")
    ws.cell(row=1, column=6, value="Period filter options (auto)")
    ws.cell(row=1, column=7, value="Leavers Status normalized (auto)")
    ws.cell(row=1, column=8, value="H2: Profile headers missing flag")
    ws.cell(row=1, column=9, value="I2: Leavers headers missing flag")
    ws.cell(row=1, column=10, value="J: leaver whose Profile copy is also Inactive (auto)")
    ws.cell(row=2, column=10, value="K2 SelStart / K3 SelEnd / K4 AsOf (from FiltPeriod)")

    for r in range(R0, R1 + 1):
        ws.cell(row=r, column=1, value=(
            f'=IF(INDEX(L_ID,ROW()-4)="","",'
            f'IF(COUNTIF(P_ID,INDEX(L_ID,ROW()-4))>0,1,0))'))
        ws.cell(row=r, column=2, value=(
            f'=IF(INDEX(P_Type,ROW()-4)="","",'
            f'TRIM(UPPER(INDEX(P_Type,ROW()-4))))'))
        ws.cell(row=r, column=4, value=(
            f'=IF(INDEX(L_Type,ROW()-4)="","",'
            f'TRIM(UPPER(INDEX(L_Type,ROW()-4))))'))
        ws.cell(row=r, column=5, value=(
            f'=IF(INDEX(P_Status,ROW()-4)="","",'
            f'TRIM(UPPER(INDEX(P_Status,ROW()-4))))'))
        ws.cell(row=r, column=7, value=(
            f'=IF(INDEX(L_Status,ROW()-4)="","",'
            f'TRIM(UPPER(INDEX(L_Status,ROW()-4))))'))
        ws.cell(row=r, column=10, value=(
            f'=IF(INDEX(L_ID,ROW()-4)="","",'
            f'IF(SUMPRODUCT((P_ID=INDEX(L_ID,ROW()-4))*(P_StatusN="INACTIVE"))>0,1,0))'))

    ws.cell(row=2, column=8, value=(
        f'=IF(ISNA(MATCH(cfgStartHdr,{QDP}!$4:$4,0))'
        f'+ISNA(MATCH(cfgLWDHdr,{QDP}!$4:$4,0))'
        f'+ISNA(MATCH(cfgTypeHdr,{QDP}!$4:$4,0))'
        f'+ISNA(MATCH(cfgStatusHdr,{QDP}!$4:$4,0))'
        f'+ISNA(MATCH(cfgIDHdr,{QDP}!$4:$4,0))>0,1,0)'))
    ws.cell(row=2, column=9, value=(
        f'=IF(ISNA(MATCH(cfgStartHdr,{QDL}!$4:$4,0))'
        f'+ISNA(MATCH(cfgLWDHdr,{QDL}!$4:$4,0))'
        f'+ISNA(MATCH(cfgTypeHdr,{QDL}!$4:$4,0))'
        f'+ISNA(MATCH(cfgStatusHdr,{QDL}!$4:$4,0))'
        f'+ISNA(MATCH(cfgIDHdr,{QDL}!$4:$4,0))>0,1,0)'))

    first = (f'=IF($H$2+$I$2>0,"",IF(COUNT(P_Start)+COUNT(L_Start)=0,"",'
             f'DATE(YEAR(MIN(P_Start,L_Start)),MONTH(MIN(P_Start,L_Start)),1)))')
    ws.cell(row=5, column=3, value=first)
    for r in range(6, 5 + N_MONTHS):
        ws.cell(row=r, column=3, value=f'=IF($C$5="","",EDATE(C{r-1},1))')
    for r in range(5, 5 + N_MONTHS):
        ws.cell(row=r, column=3).number_format = "yyyy-mm"

    # Period 下拉框选项：All + 各年份 + 各月份（F5:F58）
    ws.cell(row=5, column=6, value="All")
    for k in range(5):                                   # F6:F10 年份槽位
        r = 6 + k
        ws.cell(row=r, column=6, value=(
            f'=IF($C$5="","",IF(YEAR(EDATE($C$5,12*{k}))>YEAR(EDATE($C$5,{N_MONTHS - 1})),"",'
            f'YEAR(EDATE($C$5,12*{k}))))'))
    for k in range(N_MONTHS):                            # F11:F58 月份
        src = 5 + k
        ws.cell(row=11 + k, column=6, value=f'=IF($C${src}="","",TEXT($C${src},"yyyy-mm"))')

    # 期间解析：FiltPeriod -> SelStart / SelEnd / AsOf
    ws.cell(row=2, column=11, value=(
        '=IFERROR(IF(OR(FiltPeriod="All",FiltPeriod=""),$C$5,'
        'IF(LEN(FiltPeriod)=4,DATE(VALUE(FiltPeriod),1,1),'
        'DATE(VALUE(LEFT(FiltPeriod,4)),VALUE(MID(FiltPeriod,6,2)),1))),"")'))
    ws.cell(row=3, column=11, value=(
        '=IFERROR(IF(OR(FiltPeriod="All",FiltPeriod=""),"",'
        'IF(LEN(FiltPeriod)=4,DATE(VALUE(FiltPeriod),12,31),'
        'EOMONTH(DATE(VALUE(LEFT(FiltPeriod,4)),VALUE(MID(FiltPeriod,6,2)),1),0))),"")'))
    ws.cell(row=4, column=11, value='=IF($K$3="",TODAY(),$K$3)')
    for r, fmt in ((2, "yyyy-mm-dd"), (3, "yyyy-mm-dd"), (4, "yyyy-mm-dd")):
        ws.cell(row=r, column=11).number_format = fmt

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 20
    ws.column_dimensions["J"].width = 14
    ws.column_dimensions["K"].width = 12


# ============ 演示数据 ============
def _demo(idn, status, full, cn, squad, vendor, cat, level, jf, gender, start, lwd=None,
          role="Engineer"):
    """构造一条 24 列演示记录（列顺序与 PROFILE_HEADERS 一致）。"""
    return [f"E{idn:04d}", f"1B{idn:04d}", status, full, cn, "IBGT", squad, vendor,
            None, level, jf, role, gender, cat, role, None,
            start, None, "IBG", "M01", "IBGT", "PC01", "D01", lwd]


DEMO_PROFILE = [
    _demo(1,    "Active",   "Zhang Wei",   "张伟", "Design", "VendorA", "Perm", "L3", "Y", "M", date(2024, 1, 8)),
    _demo(2,    "Active",   "Li Na",       "李娜", "Test",   "VendorA", "PC",   "L2", "Y", "F", date(2024, 2, 18)),
    _demo(3,    "Inactive", "Wang Qiang",  "王强", "Design", "VendorB", "MS",   "L3", "N", "M", date(2024, 3, 4),  date(2024, 11, 29)),
    _demo(4,    "Active",   "Zhao Min",    "赵敏", "Test",   "VendorB", "Perm", "L4", "Y", "F", date(2024, 1, 22), role="Lead"),
    _demo(5,    "Inactive", "Liu Yang",    "刘洋", "Design", "VendorA", "PC",   "L2", "Y", "M", date(2024, 4, 15), date(2025, 3, 31)),
    _demo(6,    "Active",   "Chen Jing",   "陈静", "Qual",   "VendorC", "MS",   "L2", "N", "F", date(2024, 5, 6)),
    _demo(7,    "Inactive", "Yang Fan",    "杨帆", "Design", "VendorA", "Perm", "L3", "Y", "M", date(2024, 6, 10), date(2025, 6, 30)),
    _demo(8,    "Active",   "Zhou Ting",   "周婷", "Test",   "VendorB", "PC",   "L2", "Y", "F", date(2024, 7, 1)),
    _demo(9,    "Inactive", "Wu Lei",      "吴磊", "Qual",   "VendorC", "MS",   "L3", "N", "M", date(2024, 8, 19), date(2026, 2, 27)),
    _demo(10,   "Active",   "Zheng Shuai", "郑爽", "Design", "VendorA", "Perm", "L2", "Y", "M", date(2024, 9, 9)),
    _demo(11,   "Inactive", "Feng Jun",    "冯军", "Test",   "VendorB", "PC",   "L3", "Y", "M", date(2024, 11, 25), date(2025, 9, 30)),
    _demo(12,   "Active",   "Xu Qing",     "许晴", "Qual",   "VendorC", "MS",   "L2", "N", "F", date(2025, 1, 6)),
    _demo(13,   "Inactive", "He Chao",     "何超", "Design", "VendorA", "Perm", "L3", "Y", "M", date(2025, 3, 17), date(2025, 12, 31)),
    _demo(14,   "Active",   "Deng Zi",     "邓紫", "Test",   "VendorB", "PC",   "L2", "Y", "F", date(2025, 5, 5)),
    _demo(15,   "Active",   "Cao Ying",    "曹颖", "Qual",   "VendorC", "MS",   "L2", "N", "F", date(2025, 7, 14)),
    _demo(16,   "Active",   "Pan Shuai",   "潘帅", "Design", "VendorA", "Perm", "L2", "Y", "M", date(2026, 1, 11)),
    _demo(17,   "Active",   "Lin Yun",     "林允", "Test",   "VendorB", "PC",   "L2", "Y", "F", date(2026, 4, 20)),
    _demo(18,   "Inactive", "Su Mang",     "苏芒", "Qual",   "VendorC", "MS",   "L2", "N", "F", date(2026, 8, 3), date(2026, 8, 29)),
]
# Leavers 表：E0007 与 Profile 重复（测试去重）；E9001-E9003 仅在 Leavers（测试合并）；
# E9002/E9003 测试分类大小写与首尾空格归一化
DEMO_LEAVER = [
    _demo(7,    "Inactive", "Yang Fan", "杨帆", "Design", "VendorA", "Perm", "L3", "Y", "M", date(2024, 6, 10), date(2025, 6, 30)),
    _demo(9001, "Inactive", "Gao Yuan", "高远", "Qual",   "VendorC", "MS",   "L2", "N", "M", date(2024, 2, 1), date(2025, 1, 15)),
    _demo(9002, "Inactive", "Tian Yu",  "田雨", "Test",   "VendorB", "perm", "L2", "Y", "F", date(2025, 2, 3), date(2025, 10, 24)),
    _demo(9003, "Inactive", "Qin Lan",  "秦岚", "Design", "VendorA", "PC ",  "L3", "Y", "F", date(2025, 6, 2), date(2026, 5, 29)),
]


def write_demo_data(path):
    """把演示数据写入已生成的文件（作为值写入，日期为真日期）。"""
    from openpyxl import load_workbook
    wb = load_workbook(path)
    for sheet, rows in ((S_DP, DEMO_PROFILE), (S_DL, DEMO_LEAVER)):
        ws = wb[sheet]
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                cell = ws.cell(row=R0 + i, column=2 + j)
                if v is not None:
                    cell.value = v
                if j in (COL_START, COL_CED, COL_LWD):
                    cell.number_format = "yyyy-mm-dd"
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


# ============ Python 基准（验证用） ============
def _norm(s):
    return (s or "").strip().upper()


def _is_date(d):
    return hasattr(d, "year")


def ground_truth(filter_cat=None, today=None):
    """返回 (每月 dict 列表, 当前在职 dict)。与 Excel 公式口径严格一致（Status 口径）。"""
    import calendar
    today = today or date.today()
    p_rows = [{"id": r[COL_ID], "status": _norm(r[COL_STATUS]), "type": _norm(r[COL_TYPE]),
               "start": r[COL_START], "lwd": r[COL_LWD]} for r in DEMO_PROFILE]
    p_by_id = {r["id"].strip().upper(): r for r in p_rows}
    p_ids = set(p_by_id)
    l_rows = []
    for r in DEMO_LEAVER:
        key = r[COL_ID].strip().upper()
        l_rows.append({"id": r[COL_ID], "status": _norm(r[COL_STATUS]),
                       "type": _norm(r[COL_TYPE]), "start": r[COL_START], "lwd": r[COL_LWD],
                       "inp": key in p_ids,
                       "twin_ina": key in p_by_id and p_by_id[key]["status"] == "INACTIVE"})

    def match(t, cat):
        return True if cat is None else t == cat.upper()

    def emp(r, eom, cat):
        # Active 一直算在职；Inactive 算到 Last Work Day 当月末；无 LWD 的 Inactive 不计入
        if not (r["start"] and r["start"] <= eom and match(r["type"], cat)):
            return False
        if r["status"] == "ACTIVE":
            return True
        return bool(r["lwd"] and _is_date(r["lwd"]) and r["lwd"] >= eom)

    def left_in(r, lo, hi, cat):
        # Status=Inactive 且 LWD 落在 [lo, hi]；两表重复时以 Leavers 为准（Profile 副本
        # 也是 Inactive 才需要去重）
        return bool(r["status"] == "INACTIVE" and r["lwd"] and _is_date(r["lwd"])
                    and lo <= r["lwd"] <= hi and match(r["type"], cat))

    starts = [r["start"] for r in p_rows + l_rows if r["start"]]
    first = date(min(starts).year, min(starts).month, 1)
    months = []
    for k in range(N_MONTHS):
        y = first.year + (first.month - 1 + k) // 12
        mth = (first.month - 1 + k) % 12 + 1
        m_first = date(y, mth, 1)
        eom = date(y, mth, calendar.monthrange(y, mth)[1])
        nxt_first = date(y + (mth // 12), mth % 12 + 1, 1)

        row = {"month": f"{y}-{mth:02d}", "hc": 0, "out": 0, "join": 0,
               "out_perm": 0, "out_pc": 0, "out_ms": 0}
        for cat in (None, "PERM", "PC", "MS"):
            hc = (sum(emp(r, eom, cat) for r in p_rows)
                  + sum(emp(r, eom, cat) for r in l_rows)
                  - sum(emp(r, eom, cat) and r["inp"] for r in l_rows))
            out_p = sum(1 for r in p_rows if left_in(r, m_first, eom, cat))
            out_l = sum(1 for r in l_rows if left_in(r, m_first, eom, cat)
                        and not (r["inp"] and r["twin_ina"]))
            join = (sum(1 for r in p_rows if r["start"] and _is_date(r["start"])
                        and m_first <= r["start"] < nxt_first and match(r["type"], cat))
                    + sum(1 for r in l_rows if r["start"] and _is_date(r["start"])
                          and m_first <= r["start"] < nxt_first and match(r["type"], cat))
                    - sum(1 for r in l_rows if r["start"] and _is_date(r["start"])
                          and m_first <= r["start"] < nxt_first
                          and match(r["type"], cat) and r["inp"]))
            if cat is None:
                row["hc"], row["out"], row["join"] = hc, out_p + out_l, join
            else:
                row[{("PERM"): "out_perm", ("PC"): "out_pc", ("MS"): "out_ms"}[cat]] = out_p + out_l
            if filter_cat and cat == filter_cat.upper():
                row["hc"], row["out"], row["join"] = hc, out_p + out_l, join
        months.append(row)
    now = {}
    for cat in [None, "PERM", "PC", "MS"]:
        hc = (sum(1 for r in p_rows if r["status"] == "ACTIVE" and r["start"]
                  and _is_date(r["start"]) and r["start"] <= today and match(r["type"], cat))
              + sum(1 for r in l_rows if r["status"] == "ACTIVE" and r["start"]
                    and _is_date(r["start"]) and r["start"] <= today and match(r["type"], cat))
              - sum(1 for r in l_rows if r["status"] == "ACTIVE" and r["start"]
                    and _is_date(r["start"]) and r["start"] <= today
                    and match(r["type"], cat) and r["inp"]))
        now[cat or "ALL"] = hc
    return months, now


# ============ 主流程 ============
def main():
    if getattr(sys, "frozen", False):
        out_dir = os.path.dirname(sys.executable)   # 打包为 exe 后，输出到 exe 所在文件夹
    else:
        out_dir = os.path.dirname(os.path.abspath(__file__))   # 脚本方式运行，输出到脚本所在文件夹
    tpl = os.path.join(out_dir, "IBGT_Turnover_Dashboard.xlsx")
    demo = os.path.join(out_dir, "IBGT_Turnover_Dashboard_Demo.xlsx")

    wb = build_workbook()
    wb.save(tpl)
    print("已生成模板:", tpl)

    import shutil
    shutil.copyfile(tpl, demo)
    write_demo_data(demo)
    print("已生成演示:", demo)

    months, now = ground_truth()
    print("Python 基准（全部）：当前在职 =", now)
    for m in months[:6]:
        print("  ", m["month"], "期末在职", m["hc"], "流出", m["out"],
              "Perm/PC/MS", m["out_perm"], m["out_pc"], m["out_ms"], "入职", m["join"])


if __name__ == "__main__":
    main()
    if getattr(sys, "frozen", False):
        # 双击 exe 运行时窗口停留，能看到生成结果
        try:
            input("\n按回车键退出...")
        except EOFError:
            pass
