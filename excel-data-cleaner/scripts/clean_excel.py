"""
excel-data-cleaner 核心清洗脚本。

清洗流程划分为四个阶段(详见 SKILL.md):
  阶段一 · 多格式读入 (Input)        → read_input / detect_header_rows
  阶段二 · 格式清洗 (Format Clean)   → clean_dataframe
                                        ① 字段格式清洗 / ② 空值归一 / ③ 时间统一
                                        / ④ 字符串归一 / ⑤ 异常标识
  阶段三 · 可视化清洗 (Visual Clean)  → _apply_styles / _find_effective_range / _merge_header_row
  阶段四 · 统一输出 (Output)          → write_xlsx

功能:
- 读入 .csv / .xls / .xlsx
- 按 cleaning_rules.yaml 归一字段格式
- 空值归一(""/" "/"null"/"N/A" 等 → 空)
- 异常标识并写入 _ANOMALIES sheet
- 输出统一为 .xlsx,并附带 _CLEAN_REPORT 报告 sheet

用法:
    python clean_excel.py --input <file> --output <file.xlsx> [--rules <yaml>]
"""

from __future__ import annotations

import argparse
import math
import os
import re
import shutil
import sys
import tempfile
import unicodedata
import zipfile
from copy import deepcopy
from xml.etree import ElementTree as ET
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import yaml
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


SUPPORTED_SUFFIXES = {".csv", ".xls", ".xlsx"}


# ---------------------------------------------------------------------------
# 默认规则(未提供 rules 文件时使用)
# ---------------------------------------------------------------------------

DEFAULT_RULES: Dict[str, Any] = {
    "fields": {
        "id": {
            "column_keywords": [
                "id", "编号", "no", "code", "key", "uuid",
                "工号", "订单号", "流水号", "单号", "票号", "号",
            ],
            "cell_format": "@",
            "is_primary_key": True,
        },
        "datetime": {
            "column_keywords": ["datetime", "timestamp", "date", "time", "日期", "时间", "_dt", "_ts"],
            "output_format": "YYYY-MM-DD",
            "cell_format": "yyyy-mm-dd",
        },
        "month": {
            "column_keywords": ["付款月份", "交易月份", "month"],
            "output_format": "YYYY-MM",
            "cell_format": "yyyy-mm",
        },
        "long_text": {
            "column_keywords": [
                "desc", "remark", "备注", "说明", "comment",
                "address", "地址", "note", "notes",
            ],
            "max_length": 500,
        },
        # 百分比列: 列名含"费率/比例/占比/折算比例/折扣/费率折扣/使用比"
        # 的字段, 默认 cell_format 用 0.00%(不影响底层值, 只改显示)。
        # 默认认为是 0~1 的小数比例; 若实际是 0~100 的整数百分比, 需在调用前手动调整 max_value。
        # 注意: 必须在 amount 之前定义, 否则 amount.gmv 会抢先命中"结汇GMV占比"
        "percent": {
            "column_keywords": [
                "费率", "比例", "占比", "折算比例", "折扣", "费率折扣",
                "使用比", "gmv比", "percentage", "rate",
                # 更长的关键词,优先于 amount 中的 "gmv"
                "gmv占比", "费率折扣",
            ],
            "cell_format": "0.00%",
            "allow_negative": False,
            "min_value": 0,
            "max_value": 1,
        },
        "amount": {
            "column_keywords": [
                "amount", "money", "price", "金额", "价格",
                "fee", "cost", "salary", "revenue", "income",
                "gmv",  # GMV (Gross Merchandise Volume) 视为金额类
                "手续费",  # 提成核算净手续费 等
                "调整金额", "成本", "利润",
                "余额", "balance",  # 账户余额 / 可用余额 / 支付前余额
                # 更长的复合关键词,优先于 percent 中的"比例"
                "比例费金额",
            ],
            "cell_format": "0.00",
            "allow_negative": True,
        },
        "number": {
            "column_keywords": [
                "count", "qty", "quantity", "数量", "age", "score", "ratio",
                "月份数", "次数",
            ],
            "cell_format": "0.00",
        },
        "phone": {
            "column_keywords": ["phone", "mobile", "手机", "tel"],
            "cell_format": "@",
            "regex_cn": r"^\d{11}$",
        },
        "email": {
            "column_keywords": ["email", "邮箱", "mail"],
            "cell_format": "@",
            "regex": r"^[\w.+-]+@[\w-]+\.[\w.-]+$",
        },
        "boolean": {
            "column_keywords": ["is_", "flag", "enabled", "是否", "active"],
            "cell_format": "@",
        },
    },
    "empty_values": {
        "literals": [
            "", " ", "  ", "\t", "\n",
            "null", "NULL", "None", "N/A", "n/a",
            "NA", "-", "—", "--", "#N/A", "\\N",
        ],
    },
    "anomalies": {
        "future_date_tolerance_days": 1,
        "min_date": "1970-01-01",
        "out_of_range_abs": 1e12,
        "email_regex": r"^[\w.+-]+@[\w-]+\.[\w.-]+$",
        "phone_regex_cn": r"^\d{11}$",
        # 统计学异常值:超过 Mean ± N * Std 视为 STAT_OUTLIER
        "stat_outlier_std_multiplier": 2,
        "stat_outlier_min_samples": 10,
    },
    # 字符串归一
    "string_normalize": {
        "unicode_nfkc": True,            # 全角/半角、不间断空格归一
        "trim": True,                    # 字段值首尾去空白
        "strip_internal_whitespace": False,  # 文本字段不清除内部空白
        "lowercase_email": True,         # 邮箱小写
        "lowercase_untyped_text": False, # untyped 文本不做强制小写
        "collapse_internal_whitespace": False,  # 不合并内部空白
    },
}


# ---------------------------------------------------------------------------
# 规则加载
# ---------------------------------------------------------------------------

def load_rules(rules_path: Optional[Path]) -> Dict[str, Any]:
    """加载 YAML 规则;若未提供则使用内置默认规则。"""
    if rules_path is None:
        return DEFAULT_RULES
    if not rules_path.exists():
        print(f"[warn] 规则文件 {rules_path} 不存在,使用默认规则", file=sys.stderr)
        return DEFAULT_RULES
    with rules_path.open("r", encoding="utf-8") as fh:
        rules = yaml.safe_load(fh) or {}
    # 浅合并默认规则
    merged = {**DEFAULT_RULES, **rules}
    merged["fields"] = {**DEFAULT_RULES["fields"], **rules.get("fields", {})}
    merged["anomalies"] = {**DEFAULT_RULES["anomalies"], **rules.get("anomalies", {})}
    return merged


# ---------------------------------------------------------------------------
# 字段类型识别
# ---------------------------------------------------------------------------

def detect_field_type(column: str, rules: Dict[str, Any]) -> Tuple[Optional[str], Dict[str, Any]]:
    """根据列名启发式识别字段类型。

    匹配规则:
    - 在所有 ftype 的关键词中,找到**最长匹配**的关键词(更具体的关键词胜出)
    - 多个 ftype 匹配同样长度时,按 fields 字典顺序优先
    - 没匹配 → 返回 None

    例: "比例费金额(USD)" 同时含"金额"(amount) 和"比例"(percent),
        都长 2 字符 → 平局,按 ftype 顺序;若 percent 在 amount 之前 → percent
        因此字段定义需把"比例费"作为更具体的 amount 关键词(单独词组)避免冲突
    """
    col_lower = str(column).lower()
    best_type: Optional[str] = None
    best_cfg: Dict[str, Any] = {}
    best_len = 0
    for ftype, fcfg in rules["fields"].items():
        keywords = [
            kw.lower()
            for kw in fcfg.get("column_keywords", [])
            if isinstance(kw, str)
        ]
        for kw in keywords:
            if kw in col_lower and len(kw) > best_len:
                # 只接受更长匹配(避免平局被字典顺序劫持)
                best_len = len(kw)
                best_type = ftype
                best_cfg = fcfg
    return best_type, best_cfg


def detect_header_rows(
    file_path: Path,
    sheet_name: str,
    max_scan: int = 10,
) -> int:
    """自动识别表头行数。

    扫描 .csv/.xls/.xlsx 文件指定 sheet 的前 `max_scan` 行,识别"标题行"。

    算法(同列连续两行相近 → 数据起点):
      - 按列从左到右扫描,每列从上到下
      - 若找到首个同一列连续两行都非空且内容"相近"(都是日期/都是数字/都是会员号/都是邮箱等),
        则这两行之上的所有行识别为表头行
      - 跳过完全为空的行和说明/备注类文字行
      - 若扫描结束未找到,默认返回 1(常见单行表头)

    返回:表头行数 N(>= 1)
    """
    if not file_path.exists():
        return 1
    suffix = file_path.suffix.lower()
    if suffix == ".xlsx":
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)
            if sheet_name not in wb.sheetnames:
                wb.close()
                return 1
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True))
            wb.close()
            return _detect_header_rows_from_rows(rows, max_scan=max_scan)
        except Exception:
            return 1
    if suffix == ".xls":
        try:
            df_raw = pd.read_excel(file_path, header=None, nrows=max_scan, dtype=str)
            rows = [list(r) for r in df_raw.itertuples(index=False)]
            return _detect_header_rows_from_rows(rows, max_scan=max_scan)
        except Exception:
            return 1
    if suffix == ".csv":
        try:
            import csv as _csv
            with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
                reader = _csv.reader(f)
                rows = []
                for i, row in enumerate(reader):
                    if i >= max_scan:
                        break
                    rows.append(row)
            return _detect_header_rows_from_rows(rows, max_scan=max_scan)
        except UnicodeDecodeError:
            try:
                import csv as _csv
                with open(file_path, "r", encoding="gbk", newline="") as f:
                    reader = _csv.reader(f)
                    rows = []
                    for i, row in enumerate(reader):
                        if i >= max_scan:
                            break
                        rows.append(row)
                return _detect_header_rows_from_rows(rows, max_scan=max_scan)
            except Exception:
                return 1
        except Exception:
            return 1
    # 其他格式默认 1
    return 1


def _is_empty(value: Any) -> bool:
    """空值判定(对齐 excel-style-cleaner):

    - None
    - 空字符串 / 纯空白
    - 大小写不敏感的 "none" / "None" / "null" / "NULL"
    """
    if value is None:
        return True
    if isinstance(value, str):
        s = value.strip()
        if s == "":
            return True
        if s.lower() == "none":
            return True
    return False


def _get_cell_value(ws, row: int, col: int) -> Any:
    """获取单元格值(对齐 excel-style-cleaner):

    若该单元格已属于某个 merged range,返回左上角主单元格的值。
    """
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            return ws.cell(row=mr.min_row, column=mr.min_col).value
    return ws.cell(row=row, column=col).value


def _is_date_like(v: Any) -> bool:
    """粗略判定值是否是日期(用于"相近"判定)。"""
    if isinstance(v, datetime):
        return True
    if isinstance(v, date):
        return True
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return False
        # 常见日期形式
        if re.fullmatch(r"\d{4}-\d{1,2}(-\d{1,2})?", s):
            return True
        if re.fullmatch(r"\d{4}/\d{1,2}(/\d{1,2})?", s):
            return True
        if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{2,4}", s):
            return True
    return False


def _is_numeric_value_for_header(v: Any) -> bool:
    """判定 v 是否算"含数字"(用于表头行计数,日期不算、bool 不算)。"""
    if isinstance(v, bool):
        return False
    if isinstance(v, datetime):
        return False
    if isinstance(v, date):
        return False
    if isinstance(v, (int, float)):
        if isinstance(v, float) and math.isnan(v):
            return False
        return True
    return False


def _looks_like_header_row(row_vals: List[Any]) -> bool:
    """判断一行是否像表头(非说明/备注类文字)。"""
    if all(_is_empty(v) for v in row_vals):
        return False
    # 含说明/备注关键字的视为说明行,不算表头
    for v in row_vals:
        if isinstance(v, str):
            for kw in ("请", "打开", "查看", "说明", "备注", "测试", "例子"):
                if kw in v:
                    return False
    # 平均长度过长 → 视为说明文字
    strs = [v for v in row_vals if isinstance(v, str) and not _is_empty(v)]
    if strs:
        avg_len = sum(len(s) for s in strs) / len(strs)
        if avg_len > 30:
            return False
    return True


def _values_are_similar(a: Any, b: Any) -> bool:
    """判断两个值是否"相近"(同类型 / 同格式 / 同规律)。

    用于表头识别:若同一列连续两行"相近",则数据区可能从这里开始。
    """
    if _is_empty(a) or _is_empty(b):
        return False

    # 1) 都是日期 → 相近
    a_is_date = _is_date_like(a)
    b_is_date = _is_date_like(b)
    if a_is_date and b_is_date:
        return True
    if a_is_date != b_is_date:
        return False  # 一个日期一个不是 → 不相近

    # 2) 都是数字(int/float,非 bool)→ 相近
    a_is_num = _is_numeric_value_for_header(a)
    b_is_num = _is_numeric_value_for_header(b)
    if a_is_num and b_is_num:
        return True

    # 3) 都是纯数字字符串 → 相近
    a_str = str(a).strip() if a is not None else ""
    b_str = str(b).strip() if b is not None else ""
    a_digits = bool(re.fullmatch(r"-?[\d,]+(\.\d+)?%?", a_str))
    b_digits = bool(re.fullmatch(r"-?[\d,]+(\.\d+)?%?", b_str))
    if a_digits and b_digits:
        return True

    # 4) 都是"XXX号"样式 → 相近
    a_have_hao = bool(re.search(r"号$|号[-_]?\d", a_str))
    b_have_hao = bool(re.search(r"号$|号[-_]?\d", b_str))
    if a_have_hao and b_have_hao:
        return True

    # 5) 都是邮箱 → 相近
    a_email = bool(re.fullmatch(r"[\w.+-]+@[\w-]+\.[\w.-]+", a_str))
    b_email = bool(re.fullmatch(r"[\w.+-]+@[\w-]+\.[\w.-]+", b_str))
    if a_email and b_email:
        return True

    # 6) 都是 URL → 相近
    a_url = bool(re.match(r"https?://", a_str))
    b_url = bool(re.match(r"https?://", b_str))
    if a_url and b_url:
        return True

    # 7) 都是手机号(11 位)→ 相近
    a_phone = bool(re.fullmatch(r"1\d{10}", a_str))
    b_phone = bool(re.fullmatch(r"1\d{10}", b_str))
    if a_phone and b_phone:
        return True

    # 8) 弱相似:都是字符串,长度差 ≤ 2,且都含有字母数字
    if a_str and b_str:
        if abs(len(a_str) - len(b_str)) <= 2:
            if re.search(r"[A-Za-z0-9]", a_str) and re.search(r"[A-Za-z0-9]", b_str):
                return True

    return False


def _detect_header_rows_from_rows(
    rows: List[List[Any]],
    *,
    max_scan: int = 10,
) -> int:
    """从 rows(前 N 行数据)判断表头行数。

    算法(同列连续两行相近 → 数据起点):
      - 跳过全空行(不计入表头,也不参与"连续两行"判定)
      - 跳过说明/备注类文字行
      - 按列从左到右扫描,每列从上到下
      - 若找到首个同一列连续两行都非空且内容"相近",
        则这两行之上的所有有效行识别为表头行
      - 兜底返回 1
    """
    if not rows:
        return 1

    # 1) 过滤"完全空行"和"说明/备注类文字行",保留有效行号映射
    effective_rows: List[Tuple[int, List[Any]]] = []
    for idx, row in enumerate(rows):
        if idx + 1 > max_scan:
            break
        cells = list(row) if row else []
        if not _looks_like_header_row(cells):
            continue
        effective_rows.append((idx + 1, cells))

    if len(effective_rows) < 2:
        return 1

    max_col = max(len(r[1]) for r in effective_rows)

    earliest_header_count: Optional[int] = None
    for col in range(max_col):
        col_values: List[Tuple[int, Any]] = []
        for orig_row, cells in effective_rows:
            v = cells[col] if col < len(cells) else None
            if not _is_empty(v):
                col_values.append((orig_row, v))

        for i in range(len(col_values) - 1):
            row_a, val_a = col_values[i]
            row_b, val_b = col_values[i + 1]
            if _values_are_similar(val_a, val_b):
                # 找到!数据区起点 = row_a,表头行 = 该行之前的有效行数
                header_count = sum(
                    1 for orig_r, _ in effective_rows if orig_r < row_a
                )
                if header_count <= 0:
                    header_count = 1
                if earliest_header_count is None or header_count < earliest_header_count:
                    earliest_header_count = header_count
                break  # 本列已找到

        if earliest_header_count is not None:
            break

    if earliest_header_count is not None:
        return max(1, earliest_header_count)
    return 1


def _is_blank_cell(v: Any) -> bool:
    """判断单元格是否为空(等价于 None/空字符串/纯空白)。"""
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def _looks_like_number(v: Any) -> bool:
    """判断单元格值是否像数字。"""
    if isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)):
        return True
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return False
        # 数字格式(含负号、小数点、千分位、百分号)
        if re.fullmatch(r"-?[\d,]+(\.\d+)?%?", s):
            return True
    return False


# ---------------------------------------------------------------------------
# 空值归一
# ---------------------------------------------------------------------------

def normalize_empty(value: Any, empty_literals: List[str]) -> bool:
    """判断 value 是否为空值;若是,返回 True。"""
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if isinstance(value, str):
        stripped = value.strip()
        if stripped in empty_literals:
            return True
        if stripped == "":
            return True
    return False


def _normalize_string(
    value: Any,
    norm_cfg: Dict[str, Any],
    empty_literals: List[str],
) -> Any:
    """字符串归一:NFKC + trim + 空白字符替换。

    不动中间段落空白(由字段级 long_text 规则 strip 控制)。
    """
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if not isinstance(value, str):
        return value

    s = value

    # Unicode NFKC 归一(全角→半角、不间断空格 → 空格)
    if norm_cfg.get("unicode_nfkc", True):
        s = unicodedata.normalize("NFKC", s)

    # tab / 不间断空格 → 普通空格
    s = s.replace("\t", " ").replace("\xa0", " ")

    # 首尾去空白
    if norm_cfg.get("trim", True):
        s = s.strip()

    # 空字面量归一
    if s in empty_literals or s == "":
        return None

    return s


# ---------------------------------------------------------------------------
# 字段值清洗
# ---------------------------------------------------------------------------

@dataclass
class CleanResult:
    df: pd.DataFrame
    column_types: Dict[str, str] = field(default_factory=dict)
    column_formats: Dict[str, str] = field(default_factory=dict)
    anomalies: List[Dict[str, Any]] = field(default_factory=list)
    stats: Dict[str, Any] = field(default_factory=dict)


def clean_dataframe(
    df: pd.DataFrame,
    rules: Dict[str, Any],
    sheet_name: str = "Sheet1",
    header_labels: Optional[List[str]] = None,
) -> CleanResult:
    """对 DataFrame 应用所有清洗规则,返回 CleanResult。

    参数:
        df: 内容行的 DataFrame(已剥离表头行);列名是整数 0..n-1
        rules: 规则配置
        sheet_name: sheet 名(供异常明细)
        header_labels: 可选,表头行字符串列表;若提供则用此做字段类型识别,
            否则回退到 df.columns(整数)
    """
    empty_literals = rules["empty_values"]["literals"]
    anomaly_cfg = rules["anomalies"]
    norm_cfg = rules.get("string_normalize", {})

    # 1) 字符串归一(NFKC + trim + tab/nbsp 替换)
    for col in df.columns:
        df[col] = df[col].apply(
            lambda v: _normalize_string(v, norm_cfg, empty_literals)
        )

    # 2) 空值归一
    empty_count = 0
    for col in df.columns:
        df[col] = df[col].apply(
            lambda v: None if normalize_empty(v, empty_literals) else v
        )
        empty_count += df[col].isna().sum()

    # 3) 字段类型识别 + 格式清洗
    # 列名规范:用整数 0..n-1 作为 column_formats / column_types 的 key
    # (与 write_xlsx 的 column_names=list(range(n)) 对齐)。
    # 字段类型识别按 header_labels(原始表头字符串)做关键字匹配。
    column_types: Dict[Any, str] = {}
    column_formats: Dict[Any, str] = {}
    rules_hit = Counter()

    if header_labels is None:
        header_labels = [str(c) for c in df.columns]

    for col_idx in range(len(df.columns)):
        label = header_labels[col_idx] if col_idx < len(header_labels) else str(col_idx)
        ftype, fcfg = detect_field_type(label, rules)
        if ftype is None:
            column_types[col_idx] = "untyped"
            continue
        column_types[col_idx] = ftype
        column_formats[col_idx] = fcfg.get("cell_format", "")
        rules_hit[ftype] += 1

        col_name = df.columns[col_idx]
        df[col_name] = df.apply(
            lambda row, c=col_name, t=ftype, cfg=fcfg: _clean_cell(row[c], t, cfg, c, anomaly_cfg, rules_hit),
            axis=1,
        )

    # 4) 主键重复检测
    pk_col = next(
        (c for c, t in column_types.items() if t == "id"), None
    )
    if pk_col:
        _check_duplicate_ids(df, pk_col, anomaly_cfg)

    # 5) 整理异常明细
    anomalies = _collect_anomalies(df, column_types, anomaly_cfg, rules_hit, sheet_name=sheet_name)

    stats = {
        "rows": len(df),
        "columns": len(df.columns),
        "empty_normalized": int(empty_count),
        "rules_hit": dict(rules_hit),
        "anomalies_total": len(anomalies),
    }

    return CleanResult(
        df=df,
        column_types=column_types,
        column_formats=column_formats,
        anomalies=anomalies,
        stats=stats,
    )


# ---------------------------------------------------------------------------
# 单元格级清洗
# ---------------------------------------------------------------------------

def _clean_cell(
    value: Any,
    ftype: str,
    fcfg: Dict[str, Any],
    col: str,
    anomaly_cfg: Dict[str, Any],
    rules_hit: Counter,
) -> Any:
    """对单个单元格按字段类型清洗,返回清洗后的值。"""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None

    if ftype == "id":
        s = str(value).strip()
        return s if s else None

    if ftype == "long_text":
        # 只去除首尾空白(包括换行/制表符),保留段落内部空行与空格
        s = str(value).strip()
        max_len = fcfg.get("max_length", 500)
        if len(s) > max_len:
            rules_hit["text_too_long"] += 1
            return f"{s[:max_len]} [TEXT_TOO_LONG]"
        return s

    if ftype == "phone":
        digits = re.sub(r"\D", "", str(value))
        return digits if digits else None

    if ftype == "email":
        return str(value).strip().lower()

    if ftype == "boolean":
        s = str(value).strip().lower()
        if s in ("true", "1", "yes", "y", "是", "t"):
            return "True"
        if s in ("false", "0", "no", "n", "否", "f"):
            return "False"
        return s

    if ftype == "datetime":
        parsed = _parse_datetime(value)
        return parsed.strftime("%Y-%m-%d %H:%M:%S") if parsed else None

    if ftype == "month":
        parsed = _parse_month(value)
        return parsed.strftime("%Y-%m") if parsed else None

    if ftype in ("amount", "number", "percent"):
        return _parse_number(value, ftype, fcfg, anomaly_cfg)

    return value


def _parse_datetime(value: Any) -> Optional[datetime]:
    """解析日期或时间;统一输出 24h 制 datetime。

    支持:
    - YYYY-MM-DD HH:MM:SS
    - YYYY-MM-DD HH:MM
    - YYYY/MM/DD HH:MM:SS
    - DD/MM/YYYY HH:MM:SS
    - YYYY-MM-DD(纯日期,时间归零)
    - HH:MM:SS(纯时间,日期归 1970-01-01)
    - 12h 制 HH:MM:SS AM/PM(自动转 24h)
    """
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())

    s = str(value).strip()
    if not s:
        return None

    # 12h AM/PM 标记归一(强制 24h)
    s_norm = re.sub(r"\s*(AM|PM|am|pm)\s*$", "", s).strip()
    am_pm = None
    m = re.search(r"\s+(AM|PM|am|pm)\s*$", s)
    if m:
        am_pm = m.group(1).upper()

    # 完整日期+时间(优先)
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%Y.%m.%d %H:%M:%S",
    ):
        try:
            return datetime.strptime(s_norm, fmt)
        except ValueError:
            continue

    # 纯日期
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y", "%Y.%m.%d"):
        try:
            d = datetime.strptime(s_norm, fmt)
            return d
        except ValueError:
            continue

    # 纯时间(12h/24h)
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            t = datetime.strptime(s_norm, fmt)
            return t
        except ValueError:
            continue
    if am_pm:
        for fmt in ("%I:%M:%S", "%I:%M"):
            try:
                t = datetime.strptime(s_norm, fmt)
                if am_pm == "PM" and t.hour < 12:
                    t = t.replace(hour=t.hour + 12)
                if am_pm == "AM" and t.hour == 12:
                    t = t.replace(hour=0)
                return t
            except ValueError:
                continue

    # pandas 兜底
    try:
        ts = pd.to_datetime(s, errors="raise")
        if pd.isna(ts):
            return None
        return ts.to_pydatetime()
    except Exception:
        return None


def _parse_month(value: Any) -> Optional[date]:
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.replace(day=1).date() if isinstance(value, datetime) else value.replace(day=1).date()
    s = str(value).strip()
    for fmt in ("%Y-%m", "%Y/%m", "%Y.%m"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    try:
        ts = pd.to_datetime(s, errors="raise")
        return ts.replace(day=1).date()
    except Exception:
        return None


def _parse_number(
    value: Any,
    ftype: str,
    fcfg: Dict[str, Any],
    anomaly_cfg: Dict[str, Any],
) -> Optional[float]:
    s = str(value).strip()
    # 去掉货币符号、千分位
    s = re.sub(r"[¥$€£￥,]", "", s)
    try:
        num = float(s)
    except ValueError:
        return None

    allow_negative = fcfg.get("allow_negative", True)
    if not allow_negative and num < 0:
        return None

    limit_raw = anomaly_cfg.get("out_of_range_abs", 1e12)
    limit = float(limit_raw) if not isinstance(limit_raw, (int, float)) else limit_raw
    if abs(num) > limit:
        return None

    # percent 类型:额外校验 min/max_value,超出范围视为异常 → None
    # 注意:不报错,只标记为 None(避免下游误用)
    if ftype == "percent":
        min_v = fcfg.get("min_value")
        max_v = fcfg.get("max_value")
        if min_v is not None and num < float(min_v):
            return None
        if max_v is not None and num > float(max_v):
            return None
        # percent 类型不做精度截断,保留原值(下游样式层加 0.00% 显示)
        return num

    # 金额保留 2 位小数
    if ftype == "amount":
        return round(num, 2)
    return num


# ---------------------------------------------------------------------------
# 主键重复检测
# ---------------------------------------------------------------------------

def _check_duplicate_ids(df: pd.DataFrame, pk_col: str, anomaly_cfg: Dict[str, Any]) -> None:
    pass  # 异常由 _collect_anomalies 统一收集


# ---------------------------------------------------------------------------
# 异常汇总
# ---------------------------------------------------------------------------

def _collect_anomalies(
    df: pd.DataFrame,
    column_types: Dict[str, str],
    anomaly_cfg: Dict[str, Any],
    rules_hit: Optional[Counter] = None,
    sheet_name: str = "Sheet1",
) -> List[Dict[str, Any]]:
    """扫描清洗后的数据,记录异常明细(此处基于类型规则做最后一道审计)。"""
    anomalies: List[Dict[str, Any]] = []

    # 1) 主键重复
    pk_col = next((c for c, t in column_types.items() if t == "id"), None)
    if pk_col:
        seen: Dict[str, int] = {}
        for pos, (idx, val) in enumerate(df[pk_col].items()):
            row_num = pos + 2  # +1 header, +1 1-based
            if pd.isna(val):
                anomalies.append({
                    "sheet_name": sheet_name,
                    "row": row_num,
                    "column": pk_col,
                    "original_value": "",
                    "anomaly_type": "MISSING_ID",
                    "reason": "primary key is empty",
                })
                continue
            if val in seen:
                anomalies.append({
                    "sheet_name": sheet_name,
                    "row": row_num,
                    "column": pk_col,
                    "original_value": str(val),
                    "anomaly_type": "DUP_ID",
                    "reason": f"duplicate of row {seen[val] + 2}",
                })
            else:
                seen[val] = pos

    # 2) 邮箱合法性
    email_re = re.compile(anomaly_cfg.get("email_regex", r"^[\w.+-]+@[\w-]+\.[\w.-]+$"))
    for col, ftype in column_types.items():
        if ftype != "email":
            continue
        for pos, (idx, val) in enumerate(df[col].items()):
            if pd.isna(val):
                continue
            s = str(val).strip()
            if s and not email_re.match(s):
                anomalies.append({
                    "sheet_name": sheet_name,
                    "row": pos + 2,
                    "column": col,
                    "original_value": s,
                    "anomaly_type": "BAD_EMAIL",
                    "reason": "does not match email regex",
                })

    # 3) 手机号合法性(中国大陆 11 位)
    phone_re = re.compile(anomaly_cfg.get("phone_regex_cn", r"^\d{11}$"))
    for col, ftype in column_types.items():
        if ftype != "phone":
            continue
        for pos, (idx, val) in enumerate(df[col].items()):
            if pd.isna(val):
                continue
            s = str(val).strip()
            if s and not phone_re.match(s):
                anomalies.append({
                    "sheet_name": sheet_name,
                    "row": pos + 2,
                    "column": col,
                    "original_value": s,
                    "anomaly_type": "BAD_PHONE",
                    "reason": "phone must be 11 digits",
                })

    # 4) 日期/时间超出范围
    min_date = datetime.strptime(anomaly_cfg.get("min_date", "1970-01-01"), "%Y-%m-%d").date()
    future_tolerance_raw = anomaly_cfg.get("future_date_tolerance_days", 1)
    future_tolerance = int(future_tolerance_raw) if not isinstance(future_tolerance_raw, int) else future_tolerance_raw
    today_plus = date.today() + timedelta(days=future_tolerance)
    for col, ftype in column_types.items():
        if ftype != "datetime":
            continue
        for pos, (idx, val) in enumerate(df[col].items()):
            if pd.isna(val):
                continue
            try:
                dt = datetime.strptime(str(val), "%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue
            d = dt.date()
            if d < min_date or d > today_plus:
                anomalies.append({
                    "sheet_name": sheet_name,
                    "row": pos + 2,
                    "column": col,
                    "original_value": str(val),
                    "anomaly_type": "BAD_DATE",
                    "reason": f"datetime out of range [{min_date}, {today_plus}]",
                })

    # 5) 全行重复(任意一行的所有列值与之前某行完全相同)
    seen_rows: Dict[Tuple, int] = {}
    for pos, (idx, row) in enumerate(df.iterrows()):
        key = tuple(row.values)
        if key in seen_rows:
            anomalies.append({
                "sheet_name": sheet_name,
                "row": pos + 2,
                "column": "(ALL)",
                "original_value": "",
                "anomaly_type": "DUP_ROW",
                "reason": f"duplicate of row {seen_rows[key] + 2}",
            })
        else:
            seen_rows[key] = pos

    # 6) 统计学异常值(数值列:Mean ± N * Std)
    std_mult_raw = anomaly_cfg.get("stat_outlier_std_multiplier", 2)
    std_mult = float(std_mult_raw) if not isinstance(std_mult_raw, (int, float)) else std_mult_raw
    min_samples_raw = anomaly_cfg.get("stat_outlier_min_samples", 10)
    min_samples = int(min_samples_raw) if not isinstance(min_samples_raw, int) else min_samples_raw
    if rules_hit is not None:
        rules_hit["stat_outlier"] = 0
    for col, ftype in column_types.items():
        if ftype not in ("amount", "number"):
            continue
        numeric = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(numeric) < min_samples:
            continue
        mean = numeric.mean()
        std = numeric.std(ddof=0)
        if std == 0 or pd.isna(std):
            continue
        lower = mean - std_mult * std
        upper = mean + std_mult * std
        for pos, val in df[col].items():
            if pd.isna(val):
                continue
            try:
                v = float(val)
            except (TypeError, ValueError):
                continue
            if v < lower or v > upper:
                anomalies.append({
                    "sheet_name": sheet_name,
                    "row": int(pos) + 2 if not isinstance(pos, int) else pos + 2,
                    "column": col,
                    "original_value": str(val),
                    "anomaly_type": "STAT_OUTLIER",
                    "reason": f"value {v:.4f} outside [{lower:.4f}, {upper:.4f}]",
                })
                if rules_hit is not None:
                    rules_hit["stat_outlier"] += 1

    return anomalies


# ---------------------------------------------------------------------------
# 文件读写
# ---------------------------------------------------------------------------

def read_input(path: Path, sheet: Optional[str] = None) -> Dict[str, pd.DataFrame]:
    """读入任意支持格式的文件,返回 {sheet_name: DataFrame}。

    列名规范:统一使用整数 0/1/2/...(`header=None`)。
    这样原始数据(包括首行表头)都保留为内容,由 detect_header_rows
    决定哪些是表头行,避免 pandas 默认 header=0 误丢首行。
    """
    if not path.exists():
        raise FileNotFoundError(f"输入文件不存在: {path}")
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise ValueError(
            f"不支持的文件格式: {suffix};支持 {sorted(SUPPORTED_SUFFIXES)}"
        )
    if suffix == ".csv":
        df = pd.read_csv(path, header=None, dtype=str, keep_default_na=False)
        df.columns = list(range(len(df.columns)))
        return {"Sheet1": df}
    if suffix == ".xls":
        df = pd.read_excel(path, header=None, dtype=str, sheet_name=sheet).astype(str)
        if isinstance(df, dict):
            return {k: _reset_cols(v) for k, v in df.items()}
        return {sheet or "Sheet1": _reset_cols(df)}
    # .xlsx
    df = pd.read_excel(path, header=None, dtype=str, sheet_name=sheet)
    if isinstance(df, dict):
        return {k: _reset_cols(v) for k, v in df.items()}
    return {sheet or "Sheet1": _reset_cols(df)}


def _reset_cols(df: pd.DataFrame) -> pd.DataFrame:
    """把 df 的列名重置为 0..n-1。"""
    df = df.copy()
    df.columns = list(range(len(df.columns)))
    return df


def write_xlsx(
    result: CleanResult,
    output_path: Path,
    sheet_name: str = "Sheet1",
    rules: Optional[Dict[str, Any]] = None,
    header_rows_df: Optional[pd.DataFrame] = None,
    use_xlsxwriter: bool = False,
) -> None:
    """写出清洗后的 xlsx。

    本 Skill **不再支持 .xlsm 宏保留**;无论输入是什么后缀,统一输出 .xlsx。
    **不**生成 `_CLEAN_REPORT` / `_ANOMALIES` 等附加 sheet,只保留清洗后的数据 sheet。

    若提供 `header_rows_df`(由 clean_file 从原始 df 头部剥离的表头行),
    会拼回 df 头部,使输出 xlsx 完整保留"原始表头 + 清洗后的内容"。

    参数:
        use_xlsxwriter: True 时走 _write_xlsx_fast(xlsxwriter 流式写出,适合 >50 万 cells);
                       False 时走 openpyxl 默认路径(适合中小文件)
    """
    if use_xlsxwriter:
        _write_xlsx_fast(
            result=result,
            output_path=output_path,
            sheet_name=sheet_name,
            rules=rules,
            header_rows_df=header_rows_df,
        )
        return

    output_path.parent.mkdir(parents=True, exist_ok=True)
    suffix = output_path.suffix.lower()
    if suffix != ".xlsx":
        raise ValueError(f"输出文件必须是 .xlsx,当前: {suffix}")

    # 拼回表头行(如果有)+ 内容行;列名用整数 0..n-1
    df_out = result.df.copy()
    if header_rows_df is not None and len(header_rows_df) > 0:
        # 把表头列名也改成整数,便于对齐
        header_rows_df = header_rows_df.copy()
        header_rows_df.columns = list(range(len(header_rows_df.columns)))
        df_out.columns = list(range(len(df_out.columns)))
        df_out = pd.concat([header_rows_df, df_out], ignore_index=True)
    else:
        df_out.columns = list(range(len(df_out.columns)))
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, header=False, sheet_name=sheet_name)

    # 第二步:应用单元格格式(number_format),按原始列名顺序匹配(支持空列名)
    _apply_cell_formats(
        output_path, sheet_name, result.column_formats,
        column_names=list(range(len(result.df.columns))),
    )
    # 第四步:应用 _apply_styles(若启用 openpyxl 后端)
    if not use_xlsxwriter:
        _apply_styles(output_path, rules=rules, header_rows=max(1, result.stats.get("header_rows", 1)))


def _write_xlsx_fast(
    result: "CleanResult",
    output_path: Path,
    sheet_name: str = "Sheet1",
    rules: Optional[Dict[str, Any]] = None,
    header_rows_df: Optional[pd.DataFrame] = None,
) -> None:
    """大数据量场景的快速写出(xlsxwriter 后端,绕开 openpyxl wb.save() 瓶颈)。

    与 write_xlsx 的差异:
    - 写出引擎:xlsxwriter(流式写出,内存友好,适合 >50 万 cells 的文件)
    - 样式应用:在写 cell 时直接绑定 format 对象(而不是写完后再 apply_styles)
    - 表头合并:写入前扫描 header_rows,根据 _merge_header_row + _merge_header_col_cells
      规则预计算合并范围,使用 worksheet.merge_range() 一次性写出
    - 空值处理:None/NaN/"" 一律用 worksheet.write_blank(r, c, None, fmt),
      避免 xlsxwriter write() 把 None 误转为数值 0
    - 数值类型还原:对识别为 amount / number / percent 的列,纯数字字符串
      用 write_number() 写出,还原数值类型(而非文本)
    - 列宽自适应:按表头 + 内容字符数计算每列最佳宽度(text_wrap=True,数据自动换行)
    - 单元格对齐:表头左对齐(vcenter),数据居中(vcenter)
    """
    try:
        import xlsxwriter  # noqa
    except ImportError as exc:
        raise ImportError(
            "xlsxwriter 未安装;pip install xlsxwriter 后重试。"
            "或调用 write_xlsx(走 openpyxl 后端)。"
        ) from exc

    output_path.parent.mkdir(parents=True, exist_ok=True)
    suffix = output_path.suffix.lower()
    if suffix != ".xlsx":
        raise ValueError(f"输出文件必须是 .xlsx,当前: {suffix}")

    style_cfg = (rules or {}).get("output_style", {})
    font_name = style_cfg.get("font_name", "微软雅黑")
    font_size = style_cfg.get("font_size", 9)
    # xlsxwriter 期望 6 字符 RGB;传 8 字符 aRGB 会被错误地再加 'FF' 前缀变 10 字符
    # (openpyxl 解析失败);故统一剥到 6 字符
    def _to_rgb6(c: str) -> str:
        c = c.lstrip("#")
        if len(c) == 8:
            return c[2:]   # 剥 aRGB 前缀 → RGB
        return c
    font_color = _to_rgb6(style_cfg.get("font_color", "000000"))
    border_color = _to_rgb6(style_cfg.get("border_color", "000000"))
    border_style = style_cfg.get("border_style", "thin")
    header_fill = _to_rgb6(style_cfg.get("header_fill", "D9E1F2"))
    header_bold = style_cfg.get("header_bold", True)

    # 拼回表头行(如果有)+ 内容行
    df_out = result.df.copy()
    if header_rows_df is not None and len(header_rows_df) > 0:
        header_rows_df = header_rows_df.copy()
        header_rows_df.columns = list(range(len(header_rows_df.columns)))
        df_out.columns = list(range(len(df_out.columns)))
        df_out = pd.concat([header_rows_df, df_out], ignore_index=True)
    else:
        df_out.columns = list(range(len(df_out.columns)))

    # header_labels 用于字段类型识别(决定 cell_format)
    column_formats = result.column_formats  # {col_idx: cell_format}
    total_rows, total_cols = df_out.shape
    detected_header_rows = max(1, result.stats.get("header_rows", 1))

    workbook = xlsxwriter.Workbook(str(output_path))
    # base_fmt: 数据单元格通用样式 — 居中(vcenter),自动换行(text_wrap=True)
    base_fmt = workbook.add_format({
        "font_name": font_name,
        "font_size": font_size,
        "font_color": font_color,
        "border": 1,                  # thin 边框
        "border_color": border_color,
        "align": "center",            # 水平居中
        "valign": "vcenter",          # 垂直居中
        "text_wrap": True,            # 自动换行
    })
    # header_fmt: 表头 — 居中(vcenter),浅蓝底,自动换行
    header_fmt = workbook.add_format({
        "font_name": font_name,
        "font_size": font_size,
        "font_color": font_color,
        "bold": bool(header_bold),
        "bg_color": header_fill,
        "align": "center",          # 表头水平居中(v5)
        "valign": "vcenter",
        "border": 1,
        "border_color": border_color,
        "text_wrap": True,
    })

    # 每列的 cell_format(xlsxwriter 格式对象) — 继承 base_fmt 的居中/自动换行
    col_format_map: Dict[int, Any] = {}
    for col_idx in range(total_cols):
        fmt_str = column_formats.get(col_idx, "")
        # 在 base_fmt 上叠加 number_format / text_format
        opts = {
            "font_name": font_name,
            "font_size": font_size,
            "font_color": font_color,
            "border": 1,
            "border_color": border_color,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }
        if fmt_str == "@":
            opts["num_format"] = "@"
        elif fmt_str:
            opts["num_format"] = fmt_str
        col_format_map[col_idx] = workbook.add_format(opts)

    worksheet = workbook.add_worksheet(sheet_name)

    # === 预计算表头合并范围(在写之前完成)===
    # xlsxwriter 必须在写 cell 前确定 merge 范围(否则已写入的 cell 会被 merge 忽略)
    # 我们走 skill 的 _merge_header_row + _merge_header_col_cells 规则,
    # 但因 xlsxwriter 没有 ws.merged_cells,需要把 merged_cells 信息收集到一个独立结构
    merged_ranges: List[Tuple[int, int, int, int]] = []  # (r1, c1, r2, c2) 0-based

    # === 智能合并算法 (移植自 excel-style-cleaner v3.22+) ===
    # 策略:先做"列方向"的纵向合并(收集哪些 cell 被竖向覆盖),
    # 再做"行方向"的横向合并(在已被竖向合并的列上停止扩展避免重叠)。

    def _is_blank(v) -> bool:
        """判定单元格是否为空(兼容 None / NaN / 空字符串)"""
        if v is None:
            return True
        if isinstance(v, float) and pd.isna(v):
            return True
        return str(v) == ""

    # 构建表头数据集
    headers = []
    for r in range(detected_header_rows):
        headers.append([df_out.iat[r, c] for c in range(total_cols)])

    # 第 1 步:纵向合并 — 把每个非空源格向下延伸,直到遇到下一个非空 cell
    # 记录竖向覆盖的 cell 集合(供第 2 步使用)
    vert_merged_cells: Set[Tuple[int, int]] = set()
    for r in range(detected_header_rows):
        for c in range(total_cols):
            val = headers[r][c]
            if _is_blank(val):
                continue
            # 找下方连续空 cell 的最大行号
            end_r = r
            for next_r in range(r + 1, detected_header_rows):
                if _is_blank(headers[next_r][c]):
                    end_r = next_r
                    vert_merged_cells.add((next_r, c))
                else:
                    break
            if end_r > r:
                merged_ranges.append((r, c, end_r, c))

    # 第 2 步:横向合并 — 对每个非空且"竖向未参与合并"的源格,向右扩展
    # 关键:不能跨过已被纵向合并的列,否则与上方竖向合并范围重叠
    for r in range(detected_header_rows):
        for c in range(total_cols):
            val = headers[r][c]
            if _is_blank(val):
                continue
            # 跳过被纵向吞掉的子格(自身在 vert_merged_cells 中表示被上层 cell 纵向合并)
            if (r, c) in vert_merged_cells:
                continue
            # 向右扩展
            end_c = c
            for next_c in range(c + 1, total_cols):
                # 不能跨过已被纵向合并的列(避免与 R1/R2 的纵向合并冲突)
                if (r, next_c) in vert_merged_cells:
                    break
                if _is_blank(headers[r][next_c]):
                    end_c = next_c
                else:
                    break
            if end_c > c:
                merged_ranges.append((r, c, r, end_c))

    # 应用合并范围(必须先于 cell 写入之前;xlsxwriter 允许后写但被合并区会被忽略,故先把所有合并挂上)
    # 关键修复: 传原 cell 的值给 merge_range(),否则合并后标题会显示空
    for (r1, c1, r2, c2) in merged_ranges:
        v = df_out.iat[r1, c1]
        if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
            data_arg = ""
        else:
            data_arg = str(v) if not isinstance(v, (int, float)) else v
        worksheet.merge_range(r1, c1, r2, c2, data_arg, header_fmt)

    # === 写出表头行 ===
    for r in range(detected_header_rows):
        for c in range(total_cols):
            # 跳过已合并 cell(只写左上角)
            is_top_left = True
            for (r1, c1, r2, c2) in merged_ranges:
                if r1 == r and c1 == c:
                    if r1 != r2 or c1 != c2:
                        # 非左上角 → 跳过
                        is_top_left = False
                    break
                if r1 <= r <= r2 and c1 <= c <= c2:
                    is_top_left = False
                    break
            if not is_top_left:
                continue
            v = df_out.iat[r, c]
            if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
                worksheet.write_blank(r, c, None, header_fmt)
            elif isinstance(v, str):
                worksheet.write_string(r, c, v, header_fmt)
            else:
                worksheet.write(r, c, v, header_fmt)

    # === 写出数据行 ===
    # 列类型判定:用 result.column_types 判断该列是否需要还原为数值
    column_types = result.column_types
    for r in range(detected_header_rows, total_rows):
        for c in range(total_cols):
            v = df_out.iat[r, c]
            fmt = col_format_map.get(c, base_fmt)
            if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
                worksheet.write_blank(r, c, None, fmt)
                continue
            ftype = column_types.get(c)
            # 数值类列还原为数值类型
            if ftype in ("amount", "number", "percent"):
                try:
                    num = float(str(v).strip().replace(",", "").replace("¥", "").replace("$", ""))
                    worksheet.write_number(r, c, num, fmt)
                    continue
                except (ValueError, TypeError):
                    pass
            # boolean 列:True/False 字符串转 Excel 布尔
            if ftype == "boolean":
                s = str(v).strip()
                if s in ("True", "true", "1"):
                    worksheet.write_boolean(r, c, True, fmt)
                    continue
                if s in ("False", "false", "0"):
                    worksheet.write_boolean(r, c, False, fmt)
                    continue
            if isinstance(v, str):
                worksheet.write_string(r, c, v, fmt)
            elif isinstance(v, bool):
                worksheet.write_boolean(r, c, v, fmt)
            elif isinstance(v, (int, float)):
                worksheet.write_number(r, c, v, fmt)
            else:
                worksheet.write(r, c, v, fmt)

    # === 列宽自适应(按表头 + 内容字符数) ===
    # 中文字符按 2 字符宽度近似,ASCII 1 字符;min_width=8, max_width=40
    min_width = 8.0
    max_width = 40.0
    for c in range(total_cols):
        max_content_len = 0
        for r in range(detected_header_rows):
            v = df_out.iat[r, c]
            if v is None:
                continue
            if isinstance(v, str):
                w = sum(2 if ord(ch) > 127 else 1 for ch in v)
            else:
                w = len(str(v))
            if w > max_content_len:
                max_content_len = w
        # 抽样前 200 个数据行的内容(避免 184K 行全扫描太慢)
        sample_step = max(1, (total_rows - detected_header_rows) // 200)
        for rr in range(detected_header_rows, total_rows, sample_step):
            v = df_out.iat[rr, c]
            if v is None:
                continue
            if isinstance(v, str):
                w = sum(2 if ord(ch) > 127 else 1 for ch in v)
            else:
                w = len(str(v))
            if w > max_content_len:
                max_content_len = w
        if max_content_len == 0:
            width = min_width
        else:
            # padding = 2(单元格内容两侧留 1 字符空白),加 1 给边框
            width = max_content_len + 2 + 1
            if width > max_width:
                width = max_width
        worksheet.set_column(c, c, width)

    workbook.close()

    # 清理 AI 生成水印层(drawing_*.xml + sheet rels + Content_Types 声明)
    # xlsxwriter 在写 xlsx 时可能保留输入文件的 drawings 关系,导致 AI 生成水印被带入
    _strip_ai_drawing_layers(output_path)


def _strip_ai_drawing_layers(xlsx_path: Path) -> None:
    """从 .xlsx zip 包中移除所有 drawing_*.xml 文件及对应关系。

    应用场景:
    - 原文件含 WPS/Excel 插入的"AI 生成"水印 drawing 层
    - xlsxwriter 写出时可能保留输入文件的 drawings 关系,需要后处理清理
    - openpyxl 路径中也有此问题(虽然 openpyxl 不主动复制 drawings,但 sheet rels 可能保留)

    处理步骤:
    1. 删除所有 xl/drawings/drawing*.xml 文件
    2. 修改所有 xl/worksheets/_rels/sheet*.xml.rels,删除 drawing 关系
    3. 修改所有 xl/worksheets/sheet*.xml,删除 <drawing .../> 标签
    4. 修改 [Content_Types].xml,删除 drawing Override 声明

    容错:
    - 当目标文件被 Excel/WPS 占用时(常见情况),`os.replace` 会抛 WinError 5 拒绝访问。
      此时采用"原文件改名为备份 + 新文件使用原路径"的方式绕过文件锁。

    关键实现: 一次性打开 zip 读取所有内容,然后退出 with 块,最后再 move 文件。
    避免连续多次打开 zipfile 导致 Windows 文件句柄未释放引起的 PermissionError。
    """
    import zipfile
    import shutil
    import re
    import os

    if not xlsx_path.exists():
        return

    # === 第 1 步: 一次性读取 zip 全部内容 ===
    # 如果文件被 Excel/WPS 占用,直接读 zip 会失败;先用 shutil.copy 复制到临时位置读取
    drawings_to_remove = set()
    sheet_rels_to_modify = {}
    sheet_xml_to_modify = set()
    content_types_new = None
    all_items = {}  # {path: bytes}

    import tempfile
    tmp_dir = Path(tempfile.gettempdir())
    read_tmp = tmp_dir / (xlsx_path.stem + ".read_tmp.xlsx")
    try:
        try:
            with zipfile.ZipFile(str(xlsx_path), "r") as z:
                names = z.namelist()
                for name in names:
                    if re.match(r"xl/drawings/drawing[^\/]*\.xml$", name):
                        drawings_to_remove.add(name)
                for name in names:
                    all_items[name] = z.read(name)
        except (PermissionError, OSError):
            # 文件被占用,先复制到临时位置读取
            shutil.copy(str(xlsx_path), str(read_tmp))
            with zipfile.ZipFile(str(read_tmp), "r") as z:
                names = z.namelist()
                for name in names:
                    if re.match(r"xl/drawings/drawing[^\/]*\.xml$", name):
                        drawings_to_remove.add(name)
                for name in names:
                    all_items[name] = z.read(name)
    except Exception as exc:
        print(f"  [水印清理] 读取 zip 失败: {exc}", file=sys.stderr)
        if read_tmp.exists():
            try:
                os.remove(str(read_tmp))
            except:
                pass
        return
    finally:
        if read_tmp.exists():
            try:
                os.remove(str(read_tmp))
            except:
                pass

    if not drawings_to_remove:
        print(f"  [水印清理] 无 drawing 层需要清理")
        return

    # === 第 2 步: 处理 sheet rels 和 sheet xml ===
    for name, content_bytes in all_items.items():
        if name.startswith("xl/worksheets/_rels/") and name.endswith(".rels"):
            content = content_bytes.decode("utf-8")
            if 'relationships/drawing' in content:
                new_content = re.sub(
                    r'<Relationship[^>]*relationships/drawing[^>]*/?>',
                    "",
                    content,
                )
                sheet_rels_to_modify[name] = new_content
                sheet_xml_path = name.replace("_rels/", "").replace(".rels", "")
                sheet_xml_to_modify.add(sheet_xml_path)

    for sheet_xml_path in sheet_xml_to_modify:
        if sheet_xml_path in all_items:
            content = all_items[sheet_xml_path].decode("utf-8")
            new_content = re.sub(
                r'<drawing[^>]*r:id="[^"]*"[^>]*/?>',
                "",
                content,
            )
            new_content = re.sub(
                r'<drawing[^>]*/?>',
                "",
                new_content,
            )
            sheet_rels_to_modify[sheet_xml_path] = new_content

    # === 第 3 步: 处理 [Content_Types].xml ===
    content_types_path = "[Content_Types].xml"
    if content_types_path in all_items:
        ct_content = all_items[content_types_path].decode("utf-8")
        new_ct = re.sub(
            r'<Override[^>]*PartName="/xl/drawings/[^"]*"[^>]*/?>',
            "",
            ct_content,
        )
        if new_ct != ct_content:
            content_types_new = new_ct

    if not sheet_rels_to_modify and not drawings_to_remove and content_types_new is None:
        return

    # === 第 4 步: 写临时文件到其他目录(避开文件锁) ===
    import tempfile
    tmp_dir = Path(tempfile.gettempdir())
    tmp_path = tmp_dir / (xlsx_path.stem + ".tmp.xlsx")
    try:
        if tmp_path.exists():
            os.remove(str(tmp_path))
        with zipfile.ZipFile(str(tmp_path), "w", zipfile.ZIP_DEFLATED) as zout:
            for item_name, content_bytes in all_items.items():
                if item_name in drawings_to_remove:
                    continue
                if item_name in sheet_rels_to_modify:
                    zout.writestr(item_name, sheet_rels_to_modify[item_name])
                    continue
                if item_name == "[Content_Types].xml" and content_types_new is not None:
                    zout.writestr(item_name, content_types_new)
                    continue
                zout.writestr(item_name, content_bytes)

        # === 第 5 步: 替换原文件 ===
        # 策略:优先用 os.replace(原子),若失败(文件被 Excel/WPS 占用)则
        #       通过"删除原文件 + 移动新文件"绕过文件锁(注意:Excel 关闭后
        #       才能删除正在打开的文件;若不成功则把新文件另存为 .cleaned.xlsx 副产物)
        try:
            os.replace(str(tmp_path), str(xlsx_path))
            print(f"  [水印清理] 已移除 {len(drawings_to_remove)} 个 drawing 层 (原子替换)")
            return
        except (PermissionError, OSError) as exc_replace:
            # 文件被 Excel/WPS 占用 — 尝试先删除原文件(若 Excel 已释放句柄)
            backup_path = xlsx_path.with_suffix(xlsx_path.suffix + ".watermark_backup")
            cleaned_path = xlsx_path.with_name(xlsx_path.stem + "_cleaned.xlsx")
            try:
                # 先把原文件改名为备份(试图释放原路径锁)
                if backup_path.exists():
                    try:
                        os.remove(str(backup_path))
                    except Exception:
                        pass
                shutil.move(str(xlsx_path), str(backup_path))
                shutil.move(str(tmp_path), str(xlsx_path))
                # 删除备份
                try:
                    os.remove(str(backup_path))
                except Exception:
                    pass
                print(f"  [水印清理] 已移除 {len(drawings_to_remove)} 个 drawing 层 (绕过文件锁)")
                return
            except Exception:
                # 真正无法覆盖(Excel 占用中): 把清理后的文件保存为副产物
                if backup_path.exists() and not xlsx_path.exists():
                    # rename 成功了但 move 没成功 → 把 backup 改回原名
                    try:
                        shutil.move(str(backup_path), str(xlsx_path))
                    except Exception:
                        pass
                try:
                    if cleaned_path.exists():
                        os.remove(str(cleaned_path))
                    shutil.copy(str(tmp_path), str(cleaned_path))
                    print(f"  [水印清理] 原文件被占用,清理结果已另存为 {cleaned_path.name}")
                    print(f"  [水印清理] 提示:请关闭 Excel 后用 {cleaned_path.name} 覆盖 {xlsx_path.name}")
                except Exception as exc2:
                    print(f"  [水印清理] 完全失败: replace={exc_replace}, fallback={exc2}", file=sys.stderr)
    except Exception as exc:
        if tmp_path.exists():
            try:
                os.remove(str(tmp_path))
            except Exception:
                pass
        print(f"  [水印清理] 清理失败: {exc}", file=sys.stderr)


def _apply_cell_formats(
    xlsx_path: Path,
    sheet_name: str,
    column_formats: Dict[str, str],
    column_names: Optional[List[str]] = None,
) -> None:
    """应用 number_format 到数据列。

    优先按 column_names 的顺序匹配(第 N 列对应 column_names[N-1]);
    若 column_names 为 None,则回退到按表头单元格值匹配。
    """
    wb = load_workbook(xlsx_path)
    ws = wb[sheet_name]
    if column_names:
        # 按列序号匹配(支持空列名)
        for col_idx, col_name in enumerate(column_names, start=1):
            fmt = column_formats.get(col_name, "")
            if fmt:
                for row in ws.iter_rows(
                    min_row=2,
                    max_row=ws.max_row,
                    min_col=col_idx,
                    max_col=col_idx,
                ):
                    for cell in row:
                        cell.number_format = fmt
    else:
        for col_idx, col_name in enumerate(ws[1], start=1):
            fmt = column_formats.get(col_name.value, "")
            if fmt:
                for row in ws.iter_rows(
                    min_row=2,
                    max_row=ws.max_row,
                    min_col=col_idx,
                    max_col=col_idx,
                ):
                    for cell in row:
                        cell.number_format = fmt
    wb.save(xlsx_path)


def _find_effective_rows(ws) -> List[int]:
    """有效行 = 所有行号 减去 整行为空的行号(升序)。

    参考 excel-style-cleaner 的"整行/整列"概念:
    - 整行(1:1) = 含所有列(A 到最后一列)
    - 整行全为空 → 该行不属于有效行
    """
    empty = _find_empty_full_rows(ws)
    return [r for r in range(1, ws.max_row + 1) if r not in empty]


def _find_effective_cols(ws) -> List[int]:
    """有效列 = 所有列号 减去 整列为空的列号(升序)。

    参考 excel-style-cleaner 的"整行/整列"概念:
    - 整列(A:A) = 含所有行
    - 整列全为空 → 该列不属于有效列
    """
    empty = _find_empty_full_cols(ws)
    return [c for c in range(1, ws.max_column + 1) if c not in empty]


def _find_empty_full_rows(ws, start_row: int = 1, end_row: Optional[int] = None) -> set:
    """整行为空(1:1 全部单元格都为空)的行号集合。

    参数:
        start_row: 起始行(1-based,包含)
        end_row: 终止行(1-based,包含);None 表示到 ws.max_row
    """
    if end_row is None:
        end_row = ws.max_row
    empty: set = set()
    for r in range(start_row, end_row + 1):
        if all(_is_empty(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)):
            empty.add(r)
    return empty


def _find_empty_full_cols(ws, start_col: int = 1, end_col: Optional[int] = None) -> set:
    """整列为空(A:A 全部单元格都为空)的列号集合。

    参数:
        start_col: 起始列(1-based,包含)
        end_col: 终止列(1-based,包含);None 表示到 ws.max_column
    """
    if end_col is None:
        end_col = ws.max_column
    empty: set = set()
    for c in range(start_col, end_col + 1):
        if all(_is_empty(ws.cell(row=r, column=c).value) for r in range(1, ws.max_row + 1)):
            empty.add(c)
    return empty


def _find_effective_range(ws) -> Tuple[int, int, int, int]:
    """计算样式应用的"有效范围":(首有效行,末有效行,首有效列,末有效列)。

    术语对齐 excel-style-cleaner:
    - 有效行 = 非空行(整行 1:1 不全空)
    - 有效列 = 非空列(整列 A:A 不全空)
    - 有效范围 = 有效行的最小/最大行号 × 有效列的最小/最大列号
    """
    rows = _find_effective_rows(ws)
    cols = _find_effective_cols(ws)
    if not rows or not cols:
        return 0, 0, 0, 0
    return rows[0], rows[-1], cols[0], cols[-1]


def _auto_fit_columns(
    ws,
    max_row: int,
    max_col: int,
    header_rows: int,
    min_width: float = 8.0,
    max_width: float = 40.0,
) -> None:
    """按内容自适应列宽。

    - 仅计算表头行(前 N 行)+ 内容行的值
    - 中文字符按 2 字符宽度近似(全角字符 > 0x7F 计 2)
    - 列宽 = max(min_width, content_width + 2 padding)
    - 超出 max_width → wrap_text=True(让 Excel 自动换行),不强制截断
    """
    for c in range(1, max_col + 1):
        max_content_len = 0
        for r in range(1, max_row + 1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str):
                w = sum(2 if ord(ch) > 127 else 1 for ch in v)
            else:
                w = len(str(v))
            if w > max_content_len:
                max_content_len = w

        if max_content_len == 0:
            width = min_width
        else:
            width = min_width + max_content_len + 2
            if width > max_width:
                width = max_width
        ws.column_dimensions[get_column_letter(c)].width = width


def _auto_fit_row_heights(
    ws,
    max_row: int,
    max_col: int,
    header_rows: int,
    min_height: float = 15.0,
    base_line_height: float = 15.0,
    char_per_line: float = 1.0,
) -> None:
    """按内容自适应行高。

    - 范围:[1, max_row]
    - 行高(磅)≈ base_line_height × max(lines_needed_in_row)
      其中 lines_needed_in_row = max(各列换行后的行数)
    - 单列换行后行数估算: ceil(content_width / col_width)
      其中 content_width 用中文字符 2 / ASCII 1 估算
    - 行高下限:min_height(默认 15 磅)
    """
    from math import ceil
    for r in range(1, max_row + 1):
        max_lines = 1
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str):
                content_w = sum(2 if ord(ch) > 127 else 1 for ch in v)
            else:
                content_w = len(str(v))
            if content_w == 0:
                continue
            col_letter = get_column_letter(c)
            col_width = ws.column_dimensions[col_letter].width or 8.0
            lines = ceil(content_w / max(col_width - 2, 1))
            if lines > max_lines:
                max_lines = lines
        height = max(min_height, base_line_height * max_lines * char_per_line)
        ws.row_dimensions[r].height = height


def _apply_styles(
    xlsx_path: Path,
    rules: Optional[Dict[str, Any]] = None,
    header_rows: int = 1,
) -> None:
    """应用统一样式:微软雅黑、9 号字体、黑色细边框 + 表头合并。

    核心范围(对齐 excel-style-cleaner):
    - 样式与边框仅作用于"有效范围"内:
        first_row = 首个非空行号
        last_row  = 末个非空行号
        first_col = 首个非空列号
        last_col  = 末个非空列号
    - 边框规则:有效范围内所有单元格都画黑色细边框(包括范围内但值为空的)
    - 表头行:有效范围内前 N 行(N = header_rows)
    - 表头合并:仅横向(向右合并),不再做纵向合并

    性能优化(v2,2026-09-11):
    - 注册 NamedStyle,所有 cell 共享同一字体/边框对象引用,大幅减少对象创建
    - 仅遍历"有效范围"内实际存在的 cell(_cells 字典),跳过 None 占位
    - 大数据量(>10 万行)时不再使用逐单元格 for 循环,改用行级批量赋值
    """
    from openpyxl.styles import NamedStyle
    from openpyxl.workbook.defined_name import DefinedName  # noqa

    wb = load_workbook(xlsx_path)

    # 在 workbook 级别注册 NamedStyle,所有 cell 共享同一对象引用(关键优化点)
    STYLE_NAME_DATA = "_ec_data_style"
    STYLE_NAME_HEADER = "_ec_header_style"
    STYLE_NAME_REPORT_DATA = "_ec_report_data_style"

    thin = Side(border_style="thin", color="FF000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font = Font(name="微软雅黑", size=9, color="FF000000")
    header_fill = PatternFill(start_color="FFD9E1F2", end_color="FFD9E1F2", fill_type="solid")
    header_font = Font(name="微软雅黑", size=9, bold=True, color="FF000000")
    header_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    data_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 数据 cell 样式:边框 + 字体 + 自动换行
    data_style = NamedStyle(name=STYLE_NAME_DATA)
    data_style.font = font
    data_style.border = border
    data_style.alignment = data_align

    # 表头 cell 样式:边框 + 加粗字体 + 浅蓝底 + 左中对齐 + 自动换行
    header_style = NamedStyle(name=STYLE_NAME_HEADER)
    header_style.font = header_font
    header_style.border = border
    header_style.fill = header_fill
    header_style.alignment = header_align

    # 报告 sheet 样式:边框 + 字体 + 自动换行(无表头加粗)
    report_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    report_style = NamedStyle(name=STYLE_NAME_REPORT_DATA)
    report_style.font = font
    report_style.border = border
    report_style.alignment = report_align

    for nm in (STYLE_NAME_DATA, STYLE_NAME_HEADER, STYLE_NAME_REPORT_DATA):
        if nm not in wb.named_styles:
            wb.add_named_style(data_style if nm == STYLE_NAME_DATA else
                               header_style if nm == STYLE_NAME_HEADER else report_style)

    header_rows = max(1, int(header_rows))

    for ws in wb.worksheets:
        first_row, last_row, first_col, last_col = _find_effective_range(ws)
        if first_row == 0:
            continue  # 空 sheet 跳过

        # 合并逻辑仍需 col_has_value(按"右侧列整体非空"判定)
        col_has_value = [False] * (last_col + 2)
        row_has_value = [False] * (last_row + 2)
        for c in range(first_col, last_col + 1):
            col_has_value[c] = True
        for r in range(first_row, last_row + 1):
            row_has_value[r] = True

        is_data_sheet = not ws.title.startswith("_")
        if is_data_sheet:
            header_end_row = min(first_row + header_rows - 1, last_row)

            # === 高效样式应用(关键优化) ===
            # 1) 给"有效范围"内所有 cell 注册 NamedStyle(共享对象引用)
            # 2) 整列注册 column_dimensions[letter].style,让范围内的 cell 默认继承
            # 3) 仅对表头行额外覆写为 header_style
            #
            # 实现:
            # - 先用 iter_rows(min_col, max_col)遍历范围,但只访问实际有值的 cell
            #   (openpyxl 的 _cells 字典会自动跳过 None,内存友好)
            # - 每个 cell 设置 .style = "named_style_name"(字符串引用,共享对象)
            # - 表头 cell 多一步 .style 覆盖为 header_style

            # 大数据量场景(>= 50000 行)走快速路径
            total_cells_in_range = (last_row - first_row + 1) * (last_col - first_col + 1)

            if total_cells_in_range >= 500000:
                # === 快速路径:用 ws._cells 直接遍历存在的 cell ===
                # openpyxl 的 _cells 字典只包含实际有值的 cell,跳过 None
                # 大数据量场景可减少遍历次数 50-90%
                data_cell_count = 0
                header_cell_count = 0
                # 先用 named_style 默认应用到整列(column_dimensions)
                from openpyxl.utils import get_column_letter
                for c in range(first_col, last_col + 1):
                    letter = get_column_letter(c)
                    # 对范围内的列设默认样式(只对新写入有效;已存在的 cell 仍需遍历)
                    if letter in ws.column_dimensions:
                        ws.column_dimensions[letter].style = STYLE_NAME_DATA

                # 遍历实际存在的 cell
                for (row, col), cell in ws._cells.items():
                    if row < first_row or row > last_row:
                        continue
                    if col < first_col or col > last_col:
                        continue
                    # 保存已有的 number_format(由 _apply_cell_formats 设置),
                    # 防止 NamedStyle 覆写为 'General'
                    saved_fmt = cell.number_format
                    if row <= header_end_row:
                        cell.style = STYLE_NAME_HEADER
                        header_cell_count += 1
                    else:
                        cell.style = STYLE_NAME_DATA
                        data_cell_count += 1
                    if saved_fmt and saved_fmt != 'General':
                        cell.number_format = saved_fmt
                print(f"  [快速路径] {ws.title}: 表头 {header_cell_count} cells, 数据 {data_cell_count} cells")
            else:
                # === 标准路径:逐行遍历,适合小文件 ===
                for r in range(first_row, last_row + 1):
                    for c in range(first_col, last_col + 1):
                        cell = ws.cell(row=r, column=c)
                        # 保存已有的 number_format,防止 NamedStyle 覆写
                        saved_fmt = cell.number_format
                        if r <= header_end_row:
                            cell.style = STYLE_NAME_HEADER
                        else:
                            cell.style = STYLE_NAME_DATA
                        if saved_fmt and saved_fmt != 'General':
                            cell.number_format = saved_fmt

            # 4) 表头行横向合并(只在数据 sheet 中做)
            for r in range(first_row, header_end_row + 1):
                _merge_header_row(ws, r, last_row, last_col, col_has_value, row_has_value)

            # 4.5) 表头列纵向合并(对齐 excel-style-cleaner)
            # - 仅对第一列(A 列)做纵向合并
            # - 必须在横向合并完成后调用,因为纵向合并会跳过"已在合并范围内的 cell"
            _merge_header_col_cells(ws, 1, last_row, last_col)

            # 5) 列宽按内容自适应
            _auto_fit_columns(ws, last_row, last_col, header_end_row)

            # 6) 行高按内容自适应
            _auto_fit_row_heights(ws, last_row, last_col, header_end_row)
        else:
            # _CLEAN_REPORT / _ANOMALIES 仅应用基础样式
            for r in range(first_row, last_row + 1):
                for c in range(first_col, last_col + 1):
                    cell = ws.cell(row=r, column=c)
                    saved_fmt = cell.number_format
                    cell.style = STYLE_NAME_REPORT_DATA
                    if saved_fmt and saved_fmt != 'General':
                        cell.number_format = saved_fmt

    wb.save(xlsx_path)

    # 7) 防御性清理 drawing 层("AI 生成"水印的载体,删 xl/drawings/drawing*.xml + rels + Content_Types)
    _strip_ai_drawing_layers(xlsx_path)


def _merge_header_cells(
    ws,
    header_row: int,
    max_col: int,
    max_row: int,
) -> None:
    """表头行智能横向合并(从 excel-style-cleaner `merge_header_cells` 复制)。

    规则:
    - 找到该行每个非空 cell,若其右侧 cell 为空但**右侧列整体非空**
      (即该列在其它行有内容),则向右合并,直到碰到非空 cell 或整列空。
    - 支持多格合并(如 A_row:C_row)。
    """
    for c in range(1, max_col):
        cell_val = _get_cell_value(ws, header_row, c)
        if _is_empty(cell_val):
            continue
        next_c = c + 1
        next_val = _get_cell_value(ws, header_row, next_c)
        if _is_empty(next_val):
            # 检查 next_c 列在 (header_row 之外) 是否有内容
            col_has_data = False
            for r in range(1, max_row + 1):
                if r == header_row:
                    continue
                if not _is_empty(_get_cell_value(ws, r, next_c)):
                    col_has_data = True
                    break
            if col_has_data:
                end_c = next_c
                while end_c < max_col:
                    if not _is_empty(_get_cell_value(ws, header_row, end_c + 1)):
                        break
                    # 再检查 end_c+1 列
                    next_col_has_data = False
                    for r in range(1, max_row + 1):
                        if r == header_row:
                            continue
                        if not _is_empty(_get_cell_value(ws, r, end_c + 1)):
                            next_col_has_data = True
                            break
                    if next_col_has_data:
                        end_c += 1
                    else:
                        break
                merge_range = (
                    f"{get_column_letter(c)}{header_row}:"
                    f"{get_column_letter(end_c)}{header_row}"
                )
                try:
                    ws.merge_cells(merge_range)
                except Exception:
                    pass


def _merge_header_col_cells(
    ws,
    header_col: int,
    max_row: int,
    max_col: int,
) -> None:
    """表头列智能纵向合并(从 excel-style-cleaner `merge_header_col_cells` 复制)。

    - 仅对**第 1 列(A 列)**做纵向合并
    - 找到该列每个非空 cell,若其下方 cell 为空但**下方行整体非空**
      (即该行在其它列有内容),则向下合并,直到碰到非空 cell 或整行空
    - 关键检查:纵向合并前,先检查下方单元格是否已在任何合并范围内;
      若是 → 跳过本次纵向合并(已合并的优先,避免重复)
    """
    def is_in_any_merge(row: int, col: int) -> bool:
        for m in ws.merged_cells.ranges:
            if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
                return True
        return False

    for r in range(1, max_row):
        cell_val = _get_cell_value(ws, r, header_col)
        if _is_empty(cell_val):
            continue
        next_r = r + 1
        # 关键检查:下方单元格已被合并 → 跳过纵向合并
        if is_in_any_merge(next_r, header_col):
            continue
        next_val = _get_cell_value(ws, next_r, header_col)
        if _is_empty(next_val):
            row_has_data = False
            for c in range(1, max_col + 1):
                if c == header_col:
                    continue
                if not _is_empty(_get_cell_value(ws, next_r, c)):
                    row_has_data = True
                    break
            if row_has_data:
                end_r = next_r
                while end_r < max_row:
                    if is_in_any_merge(end_r + 1, header_col):
                        break
                    if not _is_empty(_get_cell_value(ws, end_r + 1, header_col)):
                        break
                    next_row_has_data = False
                    for c in range(1, max_col + 1):
                        if c == header_col:
                            continue
                        if not _is_empty(_get_cell_value(ws, end_r + 1, c)):
                            next_row_has_data = True
                            break
                    if next_row_has_data:
                        end_r += 1
                    else:
                        break
                if not is_in_any_merge(end_r, header_col):
                    merge_range = (
                        f"{get_column_letter(header_col)}{r}:"
                        f"{get_column_letter(header_col)}{end_r}"
                    )
                    try:
                        ws.merge_cells(merge_range)
                    except Exception:
                        pass


def _merge_header_row(
    ws,
    row: int,
    max_row: int,
    max_col: int,
    col_has_value: List[bool],
    row_has_value: List[bool],
) -> None:
    """表头行(第 `row` 行)合并规则。

    规则分两层:

    1. **顶层合并(top-level)** —— 若该行在 [first_col, last_col] 范围内
       **恰好有 1 个非空 cell 且其它 cell 全为空**,则将该 cell
       横向合并到整个数据矩形范围 [first_col, last_col](即"总标题行"
       跨多列展示)。

    2. **相邻空合并**(复用 _merge_header_cells 完整规则,含多格合并)

    注:本函数仅作顶层合并的特殊逻辑;常规相邻空合并由
    `_merge_header_cells` 完整实现。
    """
    if max_row < row or max_col < 1:
        return

    # 找出该行所有非空 cell 的列号
    non_blank_cols: List[int] = []
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        if cell.value is not None and cell.value != "":
            non_blank_cols.append(c)

    if not non_blank_cols:
        return

    # 规则 1:顶层合并 —— 整行只有 1 个非空 cell → 合并到所有列
    if len(non_blank_cols) == 1:
        anchor = non_blank_cols[0]
        if anchor < max_col:
            try:
                ws.merge_cells(
                    start_row=row, start_column=anchor,
                    end_row=row, end_column=max_col,
                )
            except Exception:
                pass
            return

    # 规则 2:常规相邻空合并(走 excel-style-cleaner 完整规则)
    _merge_header_cells(ws, row, max_col, max_row)



def clean_file(
    input_path: str | Path,
    output_path: str | Path,
    rules_path: Optional[str | Path] = None,
    sheet: Optional[str] = None,
    engine: str = "auto",
) -> Dict[str, Any]:
    """清洗单个文件;返回摘要字典。

    参数:
        input_path: 输入文件(.csv/.xls/.xlsx)
        output_path: 输出文件路径;必须以 .xlsx 结尾
        rules_path: 规则 YAML 路径(可选)
        sheet: 仅清洗指定的 sheet 名
        engine: 写出引擎选择
            - "auto": 自动选择(数据 >50 万 cells 走 xlsxwriter,否则 openpyxl)
            - "openpyxl": 强制走 openpyxl(适合中小文件,样式丰富)
            - "xlsxwriter": 强制走 xlsxwriter(适合大文件,流式写出)

    本 Skill **不再支持 .xlsm 宏保留**;无论输入是什么后缀,统一输出 .xlsx。
    """
    in_path = Path(input_path)
    out_path = Path(output_path)
    rules = load_rules(Path(rules_path) if rules_path else None)

    # 不保留宏时统一 .xlsx
    if out_path.suffix.lower() != ".xlsx":
        out_path = out_path.with_suffix(".xlsx")

    sheets = read_input(in_path, sheet=sheet)
    if not sheets:
        raise ValueError(f"输入文件 {in_path} 不包含任何 sheet")

    first_sheet_name = next(iter(sheets))
    raw_df = sheets[first_sheet_name].copy()

    detected_header_rows = detect_header_rows(in_path, first_sheet_name)

    # 阶段二清洗仅作用于"内容行"(剥离表头行),保证表头字符串不被数值/日期解析误改。
    # 表头字符串(R1..R[detected_header_rows-1])作为字段识别用列名,传给 clean_dataframe。
    header_rows_df = raw_df.iloc[:detected_header_rows].reset_index(drop=True).copy()
    content_df = raw_df.iloc[detected_header_rows:].reset_index(drop=True)
    # header_labels 用最后一行表头(即最底层的字段名,如"金额"/"数量"/"客户")
    # 当多级表头时,这一层最贴近内容,字段类型识别用它最准。
    # v3 增强: 当 R3 字段为空时(多层表头中只有 R1/R2 有"组名"的情况),
    # 向下回退用 R2, 再回退用 R1,确保每个列都有非空 header_label 参与字段识别。
    # 重要: 必须用 pd.notna() 排除 numpy.nan,否则 str(nan) = "nan" 会被误判为有效标签
    if detected_header_rows > 0 and len(header_rows_df) > 0:
        header_labels = []
        for c in range(len(raw_df.columns)):
            # 自下而上找到第一个非空表头
            label = ""
            for hr in range(detected_header_rows - 1, -1, -1):
                v = header_rows_df.iloc[hr].iloc[c]
                # 注意:v 可能是 numpy.nan(float),必须用 pd.notna() 判断
                if pd.notna(v) and str(v).strip() != "":
                    label = str(v).strip()
                    break
            header_labels.append(label)
    else:
        header_labels = [str(c) for c in raw_df.columns]
    result = clean_dataframe(
        content_df, rules, sheet_name=first_sheet_name,
        header_labels=header_labels,
    )
    result.stats["header_rows"] = detected_header_rows

    # 自动选择引擎
    total_cells = len(content_df) * len(content_df.columns)
    if engine == "auto":
        use_xlsxwriter = total_cells >= 500_000
    elif engine == "xlsxwriter":
        use_xlsxwriter = True
    else:
        use_xlsxwriter = False

    write_xlsx(
        result=result,
        output_path=out_path,
        sheet_name=first_sheet_name,
        rules=rules,
        header_rows_df=header_rows_df,
        use_xlsxwriter=use_xlsxwriter,
    )

    return {
        "input": str(in_path),
        "output": str(out_path),
        "input_sheets": list(sheets.keys()),
        "cleaned_sheet": first_sheet_name,
        "summary": result.stats,
        "anomalies_sample": result.anomalies[:10],
    }


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Excel 数据清洗")
    parser.add_argument("--input", required=True, help="输入文件(.csv/.xls/.xlsx)")
    parser.add_argument("--output", required=True, help="输出文件(.xlsx)")
    parser.add_argument("--rules", help="清洗规则 YAML 文件(可选)")
    parser.add_argument("--sheet", help="仅清洗指定的 sheet 名(默认全部)")
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    try:
        result = clean_file(
            args.input,
            args.output,
            args.rules,
            sheet=args.sheet,
        )
    except (FileNotFoundError, ValueError) as exc:
        print(f"[error] {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        import traceback
        traceback.print_exc()
        print(f"[error] {exc}", file=sys.stderr)
        return 1
    print("=== Excel Data Cleaner Summary ===")
    print(f"input          : {result['input']}")
    print(f"output         : {result['output']}")
    print(f"input_sheets   : {result['input_sheets']}")
    print(f"cleaned_sheet  : {result['cleaned_sheet']}")
    print(f"header_rows    : {result['summary'].get('header_rows', 1)}")
    for k, v in result["summary"].items():
        print(f"{k:20s}: {v}")
    if result["anomalies_sample"]:
        print("\nfirst anomalies (up to 10):")
        for a in result["anomalies_sample"]:
            print(f"  {a}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
