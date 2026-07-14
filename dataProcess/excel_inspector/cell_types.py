from __future__ import annotations

import re
from datetime import date, datetime, time
from typing import Any

import numpy as np
import pandas as pd

from .utils import clean_text, is_empty_value, to_float_or_nan

_DATE_PATTERNS = [
    re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$"),
    re.compile(r"^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$"),
    re.compile(r"^\d{4}年\d{1,2}月\d{1,2}日$"),
]
_TIME_PATTERNS = [re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$")]


def infer_cell_type(v: Any) -> str:
    if is_empty_value(v):
        return "empty"
    if isinstance(v, bool):
        return "boolean"
    if isinstance(v, (datetime, pd.Timestamp)):
        return "datetime"
    if isinstance(v, date):
        return "date"
    if isinstance(v, time):
        return "time"
    if isinstance(v, (int, np.integer)):
        return "integer"
    if isinstance(v, (float, np.floating)):
        return "number"
    s = clean_text(v)
    if s.startswith("="):
        return "formula"
    if s.endswith("%") and not np.isnan(to_float_or_nan(s)):
        return "percentage"
    if not np.isnan(to_float_or_nan(s)):
        f = to_float_or_nan(s)
        return "integer" if float(f).is_integer() else "number"
    if any(p.match(s) for p in _DATE_PATTERNS):
        return "date"
    if any(p.match(s) for p in _TIME_PATTERNS):
        return "time"
    return "string"


def aggregate_type(types: list[str]) -> tuple[str, float]:
    non_empty = [t for t in types if t != "empty"]
    if not non_empty:
        return "empty", 1.0
    numeric_like = {"integer", "number", "percentage"}
    date_like = {"date", "datetime", "time"}
    counts = {t: non_empty.count(t) for t in set(non_empty)}
    best, best_count = max(counts.items(), key=lambda kv: kv[1])
    if sum(counts.get(t, 0) for t in numeric_like) / len(non_empty) >= 0.8:
        return "number", sum(counts.get(t, 0) for t in numeric_like) / len(non_empty)
    if sum(counts.get(t, 0) for t in date_like) / len(non_empty) >= 0.8:
        return "datetime", sum(counts.get(t, 0) for t in date_like) / len(non_empty)
    confidence = best_count / len(non_empty)
    if confidence < 0.65:
        return "mixed", confidence
    return best, confidence
