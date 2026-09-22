"""excel_writer.py - v3.12+：xlsxwriter 写出引擎

提供 write_styled_xlsx（单 sheet）和 write_multi_sheet_xlsx（多 sheet）两个 API。
与 openpyxl 版（excel_style_cleaner.py）共享相同的样式规则（数字格式、关键字、颜色、列宽），
但底层库换成 xlsxwriter（性能更好、纯写新表、适合 BI / 批量出表场景）。

依赖：
  pip install xlsxwriter>=3.2
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# v3.18+：xlsxwriter 写出后清空 docProps/core.xml 中的工具水印
# （默认会写 creator="python-xlsxwriter"、last_modified_by="..." 等）
_EMPTY_XLSXWRITER_PROPERTIES: Dict[str, str] = {
    "title": "",
    "subject": "",
    "author": "",
    "manager": "",
    "company": "",
    "category": "",
    "keywords": "",
    "comments": "",
    "status": "",
    "hyperlink_base": "",
    "created": "",
    "modified": "",
    "last_modified_by": "",
}


def _strip_watermarks_xlsxwriter(wb: "xlsxwriter.Workbook") -> None:
    """v3.18+：清空 xlsxwriter 写出的 docProps/core.xml 元数据。"""
    try:
        wb.set_properties(_EMPTY_XLSXWRITER_PROPERTIES)
    except Exception:
        pass


# ============================================================
# 颜色 / 字体 / 格式常量
# ============================================================
COLOR_BORDER = "#000000"
COLOR_YELLOW_BAR = "#FFD966"
COLOR_RED_BAR = "#E06666"
COLOR_GREEN_BAR = "#93C47D"
COLOR_HEADER_BG = "#DCE6F1"

FONT_NAME = "Microsoft YaHei"
FONT_SIZE = 9

# 数字格式
FMT_PERCENT = "0.00%;[Red]-0.00%"
# v3.20+：万级格式改为"底层值 /10000 + 简洁 fmt `0.0万`"（避免显示多位）
FMT_WAN = "0.0万;[Red]-0.0万"
WAN_DIVISOR = 10000
FMT_THOUSAND = "#,##0.00;[Red]-#,##0.00"
FMT_DEFAULT = "0.00"


# ============================================================
# 列类型 → 数字格式映射
# ============================================================
def get_column_format(col_type: str) -> str:
    if col_type == "rate":
        return FMT_PERCENT
    if col_type == "money":
        return FMT_WAN
    if col_type == "thousand":
        return FMT_THOUSAND
    return FMT_DEFAULT


# ============================================================
# 字符宽度估算（与 openpyxl 版共享）
# ============================================================
def _display_width(s) -> float:
    if s is None:
        return 0
    s = str(s)
    w = 0.0
    for ch in s:
        if ord(ch) > 127:
            w += 2.0
        elif ch in ".,%":
            w += 0.5
        else:
            w += 1.0
    return w


def _calc_column_width(content_max: float, header_max: float) -> float:
    return max(content_max + 2.0, (header_max + 1.0) / 2.0 + 2.0)


def _compute_widths(headers, data, column_specs):
    n = len(headers)
    widths = [10.0] * n
    for c in range(n):
        header_max = _display_width(headers[c])
        content_max = 0.0
        for row in data:
            if c < len(row):
                content_max = max(content_max, _display_width(row[c]))
        # padding：万级 +0.5、百分比 +0.5
        spec = column_specs[c] if c < len(column_specs) else {}
        pad = 2.0
        if spec.get("type") == "money":
            pad = 2.5
        elif spec.get("type") == "rate":
            pad = 2.5
        widths[c] = _calc_column_width(content_max + (pad - 2.0), header_max)
    return widths


def _is_rate_header(s: str) -> bool:
    if not s:
        return False
    # v3.23+：与 RATE_KEYWORDS 对齐，只保留"率/比"明确指标
    return any(k in s for k in [
        "达成率", "通过率", "完成率", "完成比", "达标率",
    ])


def _is_momyoy_header(s: str) -> bool:
    if not s:
        return False
    return any(k in s for k in ["环比", "同比", "MoM", "YoY"])


def _is_money_header(s: str) -> bool:
    if not s:
        return False
    return any(k in s for k in ["金额", "收入", "支出", "回款", "定价", "销售额", "营收", "利润", "GMV"])


def _should_be_rate(spec_type: str, header: str) -> bool:
    if spec_type == "rate":
        return True
    if spec_type in ("", "auto"):
        return _is_rate_header(header)
    return False


def _should_be_money(spec_type: str, header: str) -> bool:
    if spec_type == "money":
        return True
    if spec_type in ("", "auto"):
        return _is_money_header(header)
    return False


def _should_be_momyoy(spec_type: str, header: str) -> bool:
    if spec_type in ("", "auto"):
        return _is_momyoy_header(header)
    return False


# ============================================================
# 单 sheet 写入
# ============================================================
def write_styled_xlsx(
    output_path: str,
    headers: Sequence[str],
    data: Sequence[Sequence[Any]],
    column_specs: Optional[Sequence[Dict[str, Any]]] = None,
    sheet_name: str = "Sheet1",
    bold_rows: Optional[set] = None,
) -> str:
    """v3.12+：单 sheet 写出。返回 output_path。"""
    output_path = str(output_path)
    if column_specs is None:
        column_specs = [{"type": "auto"}] * len(headers)
    if bold_rows is None:
        bold_rows = set()

    wb = xlsxwriter.Workbook(output_path)
    _strip_watermarks_xlsxwriter(wb)

    ws = wb.add_worksheet(sheet_name)

    # 通用样式
    base_fmt = wb.add_format({
        "font_name": FONT_NAME,
        "font_size": FONT_SIZE,
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "border_color": COLOR_BORDER,
    })
    header_fmt = wb.add_format({
        "font_name": FONT_NAME,
        "font_size": FONT_SIZE,
        "bold": True,
        "align": "center",
        "valign": "vcenter",
        "bg_color": COLOR_HEADER_BG,
        "border": 1,
        "border_color": COLOR_BORDER,
    })
    money_fmt = wb.add_format({
        "font_name": FONT_NAME,
        "font_size": FONT_SIZE,
        "num_format": FMT_WAN,
        "align": "right",
        "valign": "vcenter",
        "border": 1,
        "border_color": COLOR_BORDER,
    })
    percent_fmt = wb.add_format({
        "font_name": FONT_NAME,
        "font_size": FONT_SIZE,
        "num_format": FMT_PERCENT,
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "border_color": COLOR_BORDER,
    })

    # 1) 表头
    for c, h in enumerate(headers):
        is_bold = (1 in bold_rows)
        fmt = header_fmt if is_bold else base_fmt
        ws.write_string(0, c, "" if h is None else str(h), fmt)

    # 2) 数据 + 条件格式
    for r_idx, row in enumerate(data, start=1):
        is_bold_row = (r_idx in bold_rows)
        for c, v in enumerate(row):
            spec_type = column_specs[c].get("type", "auto") if c < len(column_specs) else "auto"
            header_str = str(headers[c]) if c < len(headers) else ""

            cell_fmt = base_fmt
            if is_bold_row:
                cell_fmt = header_fmt

            # 类型判定
            is_money = _should_be_money(spec_type, header_str)
            is_rate = _should_be_rate(spec_type, header_str)
            is_momyoy = _should_be_momyoy(spec_type, header_str)

            # 万级（数字列均值判定 — 简化处理：列类型=auto 时只看 header）
            apply_format = None
            if is_rate or spec_type == "rate":
                apply_format = FMT_PERCENT
                cell_fmt = percent_fmt
            elif is_money or spec_type == "money":
                apply_format = FMT_WAN
                cell_fmt = money_fmt
                # v3.20+：万级底层值 /10000，显示简洁（避免 1500000.0万）
                if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v) >= WAN_DIVISOR:
                    v = v / WAN_DIVISOR
            elif is_momyoy:
                apply_format = FMT_PERCENT
                cell_fmt = percent_fmt

            # 写入值
            if v is None or v == "":
                ws.write_blank(r_idx, c, None, cell_fmt)
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.write_number(r_idx, c, v, cell_fmt)
            else:
                ws.write_string(r_idx, c, str(v), cell_fmt)

    # 3) 列宽
    widths = _compute_widths(headers, data, column_specs)
    for c, w in enumerate(widths):
        ws.set_column(c, c, min(w, 50.0))

    # 4) 行高
    ws.set_row(0, 22)
    for r_idx in range(1, len(data) + 1):
        ws.set_row(r_idx, 18)

    # 5) 条件格式（达成率 Data Bar / 环比同比 Data Bar）
    for c, h in enumerate(headers):
        header_str = str(h) if h is not None else ""
        spec_type = column_specs[c].get("type", "auto") if c < len(column_specs) else "auto"
        col_letter = xl_col_to_name(c)

        if _is_rate_header(header_str) or spec_type == "rate":
            if len(data) > 0:
                ws.conditional_format(1, c, len(data), c, {
                    "type": "data_bar",
                    "bar_color": COLOR_YELLOW_BAR,
                })
        elif _is_momyoy_header(header_str) or spec_type in ("momyoy",):
            if len(data) > 0:
                ws.conditional_format(1, c, len(data), c, {
                    "type": "data_bar",
                    "bar_color": COLOR_RED_BAR,
                })

    wb.close()
    return output_path


# ============================================================
# 多 sheet 写入
# ============================================================
def write_multi_sheet_xlsx(
    output_path: str,
    sheets: List[Dict[str, Any]],
) -> str:
    """v3.12+：多 sheet 写出。

    sheets 每项：{
        "sheet_name": str,
        "headers": list,
        "data": list[list],
        "column_specs": list[dict] (可选),
        "header_rows_data": list[list] (可选，多层表头),
        "bold_rows": set (可选),
        "merges": list (可选),
    }
    """
    output_path = str(output_path)
    wb = xlsxwriter.Workbook(output_path)
    _strip_watermarks_xlsxwriter(wb)

    for sp in sheets:
        sheet_name = sp.get("sheet_name", "Sheet")
        headers = sp.get("headers", [])
        data = sp.get("data", [])
        column_specs = sp.get("column_specs") or [{"type": "auto"}] * len(headers)
        header_rows_data = sp.get("header_rows_data")
        bold_rows = sp.get("bold_rows") or set()
        merges = sp.get("merges") or []

        ws = wb.add_worksheet(sheet_name[:31])

        base_fmt = wb.add_format({
            "font_name": FONT_NAME,
            "font_size": FONT_SIZE,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "border_color": COLOR_BORDER,
        })
        header_fmt = wb.add_format({
            "font_name": FONT_NAME,
            "font_size": FONT_SIZE,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": COLOR_HEADER_BG,
            "border": 1,
            "border_color": COLOR_BORDER,
        })
        money_fmt = wb.add_format({
            "font_name": FONT_NAME,
            "font_size": FONT_SIZE,
            "num_format": FMT_WAN,
            "align": "right",
            "valign": "vcenter",
            "border": 1,
            "border_color": COLOR_BORDER,
        })
        percent_fmt = wb.add_format({
            "font_name": FONT_NAME,
            "font_size": FONT_SIZE,
            "num_format": FMT_PERCENT,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "border_color": COLOR_BORDER,
        })

        # 多层表头
        if header_rows_data:
            for hr_idx, hr in enumerate(header_rows_data):
                for c, v in enumerate(hr):
                    ws.write_string(hr_idx, c, "" if v is None else str(v), header_fmt)
            data_offset = len(header_rows_data)
        else:
            for c, h in enumerate(headers):
                ws.write_string(0, c, "" if h is None else str(h), header_fmt)
            data_offset = 1

        # 数据
        for r_idx, row in enumerate(data, start=data_offset):
            is_bold_row = (r_idx in bold_rows)
            for c, v in enumerate(row):
                spec_type = column_specs[c].get("type", "auto") if c < len(column_specs) else "auto"
                header_str = str(headers[c]) if c < len(headers) else ""

                cell_fmt = base_fmt
                if is_bold_row:
                    cell_fmt = header_fmt

                if _should_be_rate(spec_type, header_str):
                    cell_fmt = percent_fmt
                elif _should_be_money(spec_type, header_str):
                    cell_fmt = money_fmt
                    # v3.20+：万级底层值 /10000（显示简洁，避免 1500000.0万）
                    if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v) >= WAN_DIVISOR:
                        v = v / WAN_DIVISOR
                elif _should_be_momyoy(spec_type, header_str):
                    cell_fmt = percent_fmt

                if v is None or v == "":
                    ws.write_blank(r_idx, c, None, cell_fmt)
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    ws.write_number(r_idx, c, v, cell_fmt)
                else:
                    ws.write_string(r_idx, c, str(v), cell_fmt)

        # 列宽
        widths = _compute_widths(headers, data, column_specs)
        for c, w in enumerate(widths):
            ws.set_column(c, c, min(w, 50.0))

        # 行高
        ws.set_row(0, 22)
        for r_idx in range(1, data_offset + len(data)):
            ws.set_row(r_idx, 18)

        # 合并
        for merge in merges:
            if len(merge) == 4:
                r1, c1, r2, c2 = merge
                ws.merge_range(r1, c1, r2, c2, "", header_fmt)

        # 条件格式
        for c, h in enumerate(headers):
            header_str = str(h) if h is not None else ""
            spec_type = column_specs[c].get("type", "auto") if c < len(column_specs) else "auto"
            if _is_rate_header(header_str) or spec_type == "rate":
                if len(data) > 0:
                    ws.conditional_format(data_offset, c, data_offset + len(data) - 1, c, {
                        "type": "data_bar",
                        "bar_color": COLOR_YELLOW_BAR,
                    })
            elif _is_momyoy_header(header_str) or spec_type == "momyoy":
                if len(data) > 0:
                    ws.conditional_format(data_offset, c, data_offset + len(data) - 1, c, {
                        "type": "data_bar",
                        "bar_color": COLOR_RED_BAR,
                    })

    wb.close()
    return output_path
