from __future__ import annotations

from typing import Any, Dict

import numpy as np
import pandas as pd

from .cell_types import aggregate_type, infer_cell_type
from .utils import clean_text, excel_range, is_empty_value


def infer_header_index(sheet_df: pd.DataFrame, table: Dict[str, Any], *, max_header_rows: int = 3, max_index_cols: int = 3) -> Dict[str, Any]:
    top = table["top_row"] - 1
    left = table["left_col"] - 1
    bottom = table["bottom_row"] - 1
    right = table["right_col"] - 1
    block = sheet_df.iloc[top : bottom + 1, left : right + 1].reset_index(drop=True)
    arr = block.to_numpy(dtype=object)
    h, w = arr.shape

    row_scores = []
    for r in range(min(h, max(1, max_header_rows + 2))):
        score = _header_row_score(arr, r)
        row_scores.append((r, score))
    header_count = 1 if row_scores else 0
    if len(row_scores) >= 2 and row_scores[1][1] > 0.72 and row_scores[0][1] > 0.5:
        header_count = 2
    if len(row_scores) >= 3 and row_scores[2][1] > 0.75 and min(row_scores[0][1], row_scores[1][1]) > 0.55:
        header_count = 3
    header_count = min(header_count, max_header_rows, max(0, h - 1))

    col_scores = []
    data_start = header_count
    for c in range(min(w, max(1, max_index_cols + 2))):
        score = _index_col_score(arr[data_start:, :], c) if data_start < h else 0.0
        col_scores.append((c, score))
    index_count = 0
    if col_scores and col_scores[0][1] >= 0.62:
        index_count = 1
        if len(col_scores) > 1 and col_scores[1][1] >= 0.72:
            index_count = 2
    index_count = min(index_count, max_index_cols, max(0, w - 1))

    data_top = top + header_count
    data_left = left + index_count
    data_bottom = bottom
    data_right = right
    has_data = data_top <= data_bottom and data_left <= data_right

    header_rows_1 = [top + i + 1 for i in range(header_count)]
    index_cols_1 = [left + i + 1 for i in range(index_count)]
    labels = _build_column_labels(arr, header_count, index_count)

    return {
        "table_id": table["table_id"],
        "sheet": table["sheet"],
        "header": {
            "orientation": "row",
            "rows": header_rows_1,
            "range": excel_range(top + 1, left + 1, top + header_count, right + 1) if header_count else None,
            "confidence": round(max([s for _, s in row_scores[: max(1, header_count)]] or [0]), 4),
            "scores": [{"row": top + r + 1, "score": round(s, 4)} for r, s in row_scores],
        },
        "index": {
            "orientation": "column",
            "cols": index_cols_1,
            "range": excel_range(data_top + 1, left + 1, bottom + 1, left + index_count) if index_count and data_start < h else None,
            "confidence": round(max([s for _, s in col_scores[: max(1, index_count)]] or [0]), 4),
            "scores": [{"col": left + c + 1, "score": round(s, 4)} for c, s in col_scores],
        },
        "data_region": {
            "range": excel_range(data_top + 1, data_left + 1, data_bottom + 1, data_right + 1) if has_data else None,
            "top_row": data_top + 1 if has_data else None,
            "left_col": data_left + 1 if has_data else None,
            "bottom_row": data_bottom + 1 if has_data else None,
            "right_col": data_right + 1 if has_data else None,
            "row_count": data_bottom - data_top + 1 if has_data else 0,
            "col_count": data_right - data_left + 1 if has_data else 0,
        },
        "column_labels": labels,
    }


def _header_row_score(arr: np.ndarray, r: int) -> float:
    values = list(arr[r, :])
    non_empty = [v for v in values if not is_empty_value(v)]
    if not non_empty:
        return 0.0
    string_ratio = sum(infer_cell_type(v) in {"string", "date", "datetime"} for v in non_empty) / len(non_empty)
    unique_ratio = len({clean_text(v) for v in non_empty}) / len(non_empty)
    fill_ratio = len(non_empty) / max(1, arr.shape[1])
    below_stability = 0.0
    if r + 1 < arr.shape[0]:
        stable_cols = 0
        total_cols = 0
        for c in range(arr.shape[1]):
            types = [infer_cell_type(v) for v in arr[r + 1 :, c] if not is_empty_value(v)]
            if len(types) >= 2:
                total_cols += 1
                _, conf = aggregate_type(types)
                if conf >= 0.7:
                    stable_cols += 1
        below_stability = stable_cols / total_cols if total_cols else 0.0
    return 0.30 * string_ratio + 0.20 * unique_ratio + 0.20 * fill_ratio + 0.30 * below_stability


def _index_col_score(data_arr: np.ndarray, c: int) -> float:
    if data_arr.size == 0 or c >= data_arr.shape[1]:
        return 0.0
    values = list(data_arr[:, c])
    non_empty = [v for v in values if not is_empty_value(v)]
    if not non_empty:
        return 0.0
    types = [infer_cell_type(v) for v in non_empty]
    text_or_date_ratio = sum(t in {"string", "date", "datetime", "time", "integer"} for t in types) / len(types)
    unique_ratio = len({clean_text(v) for v in non_empty}) / len(non_empty)
    fill_ratio = len(non_empty) / max(1, len(values))
    # Index columns often have fewer arithmetic-looking numeric values than data columns.
    numeric_ratio = sum(t in {"integer", "number", "percentage"} for t in types) / len(types)
    penalty = 0.75 * numeric_ratio
    return max(0.0, 0.35 * text_or_date_ratio + 0.35 * unique_ratio + 0.30 * fill_ratio - penalty)


def _build_column_labels(arr: np.ndarray, header_count: int, index_count: int) -> list[dict[str, Any]]:
    labels = []
    if header_count <= 0:
        for c in range(index_count, arr.shape[1]):
            labels.append({"relative_col": c, "label": f"col_{c + 1}"})
        return labels
    for c in range(index_count, arr.shape[1]):
        parts = [clean_text(arr[r, c]) for r in range(header_count) if not is_empty_value(arr[r, c])]
        labels.append({"relative_col": c, "label": " / ".join(parts) if parts else f"col_{c + 1}"})
    return labels
