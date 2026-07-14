from __future__ import annotations

import math
import re
from dataclasses import asdict, is_dataclass
from datetime import date, datetime, time
from decimal import Decimal
from typing import Any, Iterable

import numpy as np
import pandas as pd


def is_empty_value(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def clean_text(v: Any) -> str:
    if is_empty_value(v):
        return ""
    return str(v).strip()


def excel_col_name(n: int) -> str:
    """1-based column number to Excel column name."""
    name = ""
    while n:
        n, r = divmod(n - 1, 26)
        name = chr(65 + r) + name
    return name


def excel_cell(row_1: int, col_1: int) -> str:
    return f"{excel_col_name(col_1)}{row_1}"


def excel_range(top_row_1: int, left_col_1: int, bottom_row_1: int, right_col_1: int) -> str:
    if top_row_1 == bottom_row_1 and left_col_1 == right_col_1:
        return excel_cell(top_row_1, left_col_1)
    return f"{excel_cell(top_row_1, left_col_1)}:{excel_cell(bottom_row_1, right_col_1)}"


def safe_json(obj: Any) -> Any:
    if is_dataclass(obj):
        return safe_json(asdict(obj))
    if isinstance(obj, dict):
        return {str(k): safe_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [safe_json(x) for x in obj]
    if isinstance(obj, (datetime, date, time)):
        return obj.isoformat()
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating,)):
        if math.isnan(float(obj)):
            return None
        return float(obj)
    if isinstance(obj, np.ndarray):
        return safe_json(obj.tolist())
    if isinstance(obj, Decimal):
        return float(obj)
    if isinstance(obj, float) and math.isnan(obj):
        return None
    return obj


def to_float_or_nan(v: Any) -> float:
    if is_empty_value(v):
        return float("nan")
    if isinstance(v, bool):
        return float("nan")
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if s.endswith("%"):
            try:
                return float(s[:-1]) / 100.0
            except ValueError:
                return float("nan")
        try:
            return float(s)
        except ValueError:
            return float("nan")
    return float("nan")


def approx_equal(actual: np.ndarray, expected: np.ndarray, abs_tol: float, rel_tol: float) -> np.ndarray:
    return np.abs(actual - expected) <= np.maximum(abs_tol, np.abs(expected) * rel_tol)


def normalize_label(label: Any) -> str:
    return re.sub(r"\s+", "", clean_text(label).lower())
