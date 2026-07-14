from __future__ import annotations

from typing import Any, Dict, List, Tuple

import numpy as np
import pandas as pd

from .utils import approx_equal, excel_cell, to_float_or_nan


def detect_relations(
    sheet_df: pd.DataFrame,
    table: Dict[str, Any],
    layout: Dict[str, Any],
    *,
    abs_tol: float = 1e-6,
    rel_tol: float = 1e-4,
    min_support_count: int = 3,
    min_support_rate: float = 0.75,
    max_cols: int = 80,
    max_rows: int = 1000,
    max_k: int = 6,
    max_violations: int = 20,
) -> List[Dict[str, Any]]:
    region = layout["data_region"]
    if not region.get("range"):
        return []
    top = region["top_row"] - 1
    left = region["left_col"] - 1
    bottom = min(region["bottom_row"] - 1, top + max_rows - 1)
    right = min(region["right_col"] - 1, left + max_cols - 1)
    data = sheet_df.iloc[top : bottom + 1, left : right + 1].to_numpy(dtype=object)
    num = np.vectorize(to_float_or_nan)(data).astype(float)
    labels = _labels(layout, table)
    relations: List[Dict[str, Any]] = []

    relations.extend(_detect_axis_relations(num, axis="column", base_row=top + 1, base_col=left + 1, labels=labels, abs_tol=abs_tol, rel_tol=rel_tol, min_support_count=min_support_count, min_support_rate=min_support_rate, max_k=max_k, max_violations=max_violations))
    # Row mode transposes numeric matrix; labels are synthetic row labels.
    row_labels = {i: f"row_{top + i + 1}" for i in range(num.shape[0])}
    relations.extend(_detect_axis_relations(num.T, axis="row", base_row=top + 1, base_col=left + 1, labels=row_labels, abs_tol=abs_tol, rel_tol=rel_tol, min_support_count=min_support_count, min_support_rate=min_support_rate, max_k=max_k, max_violations=max_violations))
    relations.sort(key=lambda x: (-x["confidence"], -x["support_rate"], x["axis"], x["formula_pattern"]))
    return relations[:200]


def _labels(layout: Dict[str, Any], table: Dict[str, Any]) -> Dict[int, str]:
    table_left = table["left_col"] - 1
    data_left = layout["data_region"]["left_col"] - 1
    out = {}
    for x in layout.get("column_labels", []):
        abs_offset = table_left + x["relative_col"] - data_left
        if abs_offset >= 0:
            out[abs_offset] = x["label"]
    return out


def _detect_axis_relations(mat: np.ndarray, *, axis: str, base_row: int, base_col: int, labels: Dict[int, str], abs_tol: float, rel_tol: float, min_support_count: int, min_support_rate: float, max_k: int, max_violations: int) -> List[Dict[str, Any]]:
    # mat shape: observations x variables.
    obs, vars_ = mat.shape
    out: List[Dict[str, Any]] = []
    if obs < min_support_count or vars_ < 3:
        return out
    candidates: List[Tuple[int, int, int, int | None]] = []
    # adjacent triples
    for i in range(vars_ - 2):
        candidates.append((i, i + 1, i + 2, 1))
    # lag k triples: target is i+2k = i op i+k.
    for k in range(2, min(max_k, vars_ // 2) + 1):
        for i in range(vars_ - 2 * k):
            candidates.append((i, i + k, i + 2 * k, k))

    for a, b, t, k in candidates:
        av, bv, tv = mat[:, a], mat[:, b], mat[:, t]
        tests = [
            ("add", "+", av + bv),
            ("subtract", "-", av - bv),
            ("subtract", "-", bv - av),
            ("multiply", "*", av * bv),
        ]
        div1 = np.divide(av, bv, out=np.full_like(av, np.nan), where=~np.isclose(bv, 0, equal_nan=False))
        div2 = np.divide(bv, av, out=np.full_like(bv, np.nan), where=~np.isclose(av, 0, equal_nan=False))
        tests.extend([("divide", "/", div1), ("divide", "/", div2)])
        for rel_type, op, expected in tests:
            valid = ~np.isnan(av) & ~np.isnan(bv) & ~np.isnan(tv) & ~np.isnan(expected)
            total = int(valid.sum())
            if total < min_support_count:
                continue
            matched = approx_equal(tv[valid], expected[valid], abs_tol, rel_tol)
            support = int(matched.sum())
            rate = support / total
            if rate < min_support_rate:
                continue
            formula = _formula(labels, a, b, t, rel_type, op, expected_is_b_minus_a=(rel_type == "subtract" and np.nanmean(np.abs(expected - (bv - av))) < np.nanmean(np.abs(expected - (av - bv))) if rel_type == "subtract" else False), expected_is_b_div_a=(rel_type == "divide" and np.nanmean(np.abs(expected - div2)) <= np.nanmean(np.abs(expected - div1)) if rel_type == "divide" else False))
            violations = _violations(valid, matched, tv, expected, axis=axis, base_row=base_row, base_col=base_col, var_index=t, max_violations=max_violations)
            out.append(
                {
                    "relation_type": rel_type,
                    "axis": axis,
                    "lag_k": k,
                    "formula_pattern": formula,
                    "variables": {"left": _var_name(labels, a), "right": _var_name(labels, b), "target": _var_name(labels, t)},
                    "support_count": support,
                    "total_count": total,
                    "support_rate": round(rate, 4),
                    "confidence": round(min(0.99, rate * min(1.0, support / max(min_support_count, 20))), 4),
                    "violations": violations,
                }
            )
    # de-duplicate formula patterns
    dedup = {}
    for r in out:
        key = (r["axis"], r["formula_pattern"])
        if key not in dedup or r["confidence"] > dedup[key]["confidence"]:
            dedup[key] = r
    return list(dedup.values())


def _var_name(labels: Dict[int, str], idx: int) -> str:
    return labels.get(idx, f"var_{idx + 1}")


def _formula(labels: Dict[int, str], a: int, b: int, t: int, rel_type: str, op: str, expected_is_b_minus_a: bool = False, expected_is_b_div_a: bool = False) -> str:
    A, B, T = _var_name(labels, a), _var_name(labels, b), _var_name(labels, t)
    if expected_is_b_minus_a:
        return f"{T} = {B} - {A}"
    if expected_is_b_div_a:
        return f"{T} = {B} / {A}"
    return f"{T} = {A} {op} {B}"


def _violations(valid: np.ndarray, matched: np.ndarray, actual: np.ndarray, expected: np.ndarray, *, axis: str, base_row: int, base_col: int, var_index: int, max_violations: int) -> List[Dict[str, Any]]:
    valid_idx = np.where(valid)[0]
    bad_obs = valid_idx[~matched]
    out = []
    for obs_idx in bad_obs[:max_violations]:
        if axis == "column":
            cell = excel_cell(base_row + int(obs_idx), base_col + var_index)
        else:
            cell = excel_cell(base_row + var_index, base_col + int(obs_idx))
        out.append({"cell": cell, "expected": float(expected[obs_idx]), "actual": float(actual[obs_idx]), "diff": float(actual[obs_idx] - expected[obs_idx])})
    return out
