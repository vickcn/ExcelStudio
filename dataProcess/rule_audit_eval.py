# -*- coding: utf-8 -*-
"""
Rule audit evaluation helpers.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence, Tuple, Union
import re

from rule_audit_types import JsonDict, PassedRuleDict
from rule_audit_normalize import normalize_equation, try_split_equation_sides


def _to_matrix(values: Any) -> List[List[Any]]:
    if not isinstance(values, list):
        return []
    if values and not isinstance(values[0], list):
        return [[v] for v in values]
    return values


def _pad_matrix(matrix: List[List[Any]]) -> List[List[Any]]:
    if not matrix:
        return []
    max_cols = max((len(r) for r in matrix), default=0)
    return [list(r) + [None] * (max_cols - len(r)) for r in matrix]


def _extract_equation_sides(rule: PassedRuleDict) -> Optional[Tuple[str, str]]:
    sides = rule.get("equation_sides")
    if isinstance(sides, list) and len(sides) == 2 and all(isinstance(x, str) for x in sides):
        left = normalize_equation(sides[0])
        right = normalize_equation(sides[1])
        if left and right:
            return left, right
    eq = rule.get("equation") or rule.get("rule") or rule.get("description")
    if isinstance(eq, str):
        return try_split_equation_sides(eq)
    return None


def _normalize_expression(expr: str) -> str:
    s = str(expr).strip().replace("＝", "=")
    s = re.sub(r"\$\s+\(", "$(", s)
    s = re.sub(r"(?<!\$)\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)", r"$(\1,\2)", s)
    s = re.sub(r"\$(\d+)", r"$(\1,0)", s)
    return s


def _eval_expr(expr: str, values: Sequence[Sequence[Any]]) -> float:
    matrix = _pad_matrix(_to_matrix(values))
    rows = len(matrix)
    cols = max((len(r) for r in matrix), default=0)

    def _get_value(r: int, c: int) -> float:
        if r < 0 or r >= rows or c < 0 or c >= cols:
            raise ValueError(f"index out of range: ({r},{c})")
        val = matrix[r][c]
        if val is None:
            raise ValueError(f"index ({r},{c}) is empty")
        try:
            return float(val)
        except Exception:
            raise ValueError(f"index ({r},{c}) not numeric: {val}")

    s = _normalize_expression(expr)

    def _replace_2d(match: re.Match) -> str:
        r = int(match.group(1))
        c = int(match.group(2))
        return str(_get_value(r, c))

    numeric = re.sub(r"\$\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)", _replace_2d, s)

    def _replace_1d(match: re.Match) -> str:
        idx = int(match.group(1))
        return str(_get_value(idx, 0))

    numeric = re.sub(r"\$(\d+)", _replace_1d, numeric)

    if not re.match(r"^[\d\.\+\-\*\/\(\)\s]+$", numeric):
        raise ValueError(f"unsafe expression: {numeric}")
    try:
        return float(eval(numeric, {"__builtins__": {}}, {}))
    except Exception as exc:
        raise ValueError(f"cannot eval: {numeric}; {exc}") from exc


def _iter_window_results(details_json: JsonDict) -> List[JsonDict]:
    if isinstance(details_json, list):
        return details_json
    if isinstance(details_json, dict):
        for key in ("detailed_results", "window_results", "results"):
            val = details_json.get(key)
            if isinstance(val, list):
                return val
    return []


def _get_window_values(window_result: JsonDict) -> List[List[Any]]:
    prompt = window_result.get("prompt_data") or {}
    values = prompt.get("values")
    if values is None:
        values = prompt.get("values_matrix")
    if values is None:
        values = window_result.get("values")
    return _to_matrix(values)


def _get_window_info(window_result: JsonDict) -> JsonDict:
    info = window_result.get("window_info")
    if isinstance(info, dict):
        return info
    return {}


def evaluate_rule_on_window(
    rule: PassedRuleDict,
    window_values: Sequence[Union[int, float]],
    tolerance: float = 0.01,
) -> JsonDict:
    result: Dict[str, Any] = {
        "ok": False,
        "left_value": None,
        "right_value": None,
        "diff": None,
        "tolerance": tolerance,
        "error": None,
        "equation": rule.get("equation") or rule.get("rule") or rule.get("description"),
    }

    sides = _extract_equation_sides(rule)
    if not sides:
        result["error"] = "missing equation_sides"
        return result

    left_expr, right_expr = sides
    result["equation_sides"] = [left_expr, right_expr]

    try:
        left_val = _eval_expr(left_expr, window_values)
        right_val = _eval_expr(right_expr, window_values)
        diff = abs(left_val - right_val)
        result["left_value"] = left_val
        result["right_value"] = right_val
        result["diff"] = diff
        result["ok"] = diff <= tolerance
        if not result["ok"]:
            result["error"] = f"diff {diff:.6g} > tol {tolerance}"
    except Exception as exc:
        result["error"] = str(exc)
        result["ok"] = False
    return result


def audit_windows_with_baseline(
    details_json: JsonDict,
    baseline_rules: List[PassedRuleDict],
    *,
    row_name: Optional[str] = None,
    tolerance: float = 0.01,
    strict_row_match: bool = False,
) -> JsonDict:
    windows = _iter_window_results(details_json)
    failed_windows: List[JsonDict] = []
    total_windows = 0
    matched_windows = 0
    total_rule_checks = 0
    failed_rule_checks = 0

    for window_result in windows:
        window_info = _get_window_info(window_result)
        window_row_name = window_info.get("start_loc_row_name")
        if row_name is not None and window_row_name != row_name:
            continue

        total_windows += 1

        if row_name is not None:
            rules_to_use = baseline_rules
        else:
            rules_to_use = [
                r for r in baseline_rules
                if r.get("_row_name") in (None, "", window_row_name)
            ]
            if not rules_to_use:
                if strict_row_match:
                    continue
                rules_to_use = baseline_rules

        if not rules_to_use:
            continue

        matched_windows += 1
        values = _get_window_values(window_result)
        failed_rules = []

        for rule in rules_to_use:
            total_rule_checks += 1
            eval_result = evaluate_rule_on_window(rule, values, tolerance=tolerance)
            if not eval_result.get("ok", False):
                failed_rule_checks += 1
                failed_entry = dict(eval_result)
                failed_entry["rule_meta"] = rule
                failed_rules.append(failed_entry)

        if failed_rules:
            failed_windows.append({
                "window_id": window_info.get("window_id"),
                "start_loc_row_name": window_row_name,
                "excel_range": window_info.get("excel_range"),
                "position": window_info.get("position"),
                "failed_rules": failed_rules,
            })

    passed_rule_checks = total_rule_checks - failed_rule_checks
    metrics = {
        "total_windows": total_windows,
        "matched_windows": matched_windows,
        "total_rule_checks": total_rule_checks,
        "failed_rule_checks": failed_rule_checks,
        "passed_rule_checks": passed_rule_checks,
        "pass_rate": passed_rule_checks / total_rule_checks if total_rule_checks else 0.0,
    }
    summary = {
        "failed_windows": len(failed_windows),
    }
    return {
        "summary": summary,
        "metrics": metrics,
        "failed_windows": failed_windows,
    }


def audit_details_cells_with_baseline(
    details_json: JsonDict,
    baseline_rules: List[PassedRuleDict],
    *,
    row_name: Optional[str] = None,
    tolerance: float = 0.01,
    strict_row_match: bool = False,
) -> JsonDict:
    from rule_audit_suspect import collect_suspect_cells

    audit_report = audit_windows_with_baseline(
        details_json,
        baseline_rules,
        row_name=row_name,
        tolerance=tolerance,
        strict_row_match=strict_row_match,
    )
    failed_windows = audit_report.get("failed_windows", [])
    suspect_cells = collect_suspect_cells(failed_windows)
    return {
        "summary": audit_report.get("summary", {}),
        "metrics": audit_report.get("metrics", {}),
        "failed_windows": failed_windows,
        "suspect_cells": suspect_cells,
    }
