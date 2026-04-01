# -*- coding: utf-8 -*-
"""
Quality checks for rule outputs.
"""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple
import re

from rule_audit_types import JsonDict
from rule_audit_normalize import normalize_equation, try_split_equation_sides


def _iter_window_results(details_json: Optional[JsonDict]) -> List[JsonDict]:
    if not details_json:
        return []
    if isinstance(details_json, list):
        return details_json
    if isinstance(details_json, dict):
        for key in ("detailed_results", "window_results", "results"):
            val = details_json.get(key)
            if isinstance(val, list):
                return val
    return []


def _get_window_shape(window_result: JsonDict) -> Optional[Tuple[int, int]]:
    prompt = window_result.get("prompt_data") or {}
    values = prompt.get("values") or prompt.get("values_matrix")
    if not isinstance(values, list):
        return None
    rows = len(values)
    cols = max((len(r) for r in values if isinstance(r, list)), default=0)
    return rows, cols


def _extract_indices(exprs: List[str]) -> List[Tuple[int, int]]:
    indices: List[Tuple[int, int]] = []
    for expr in exprs:
        if not isinstance(expr, str):
            continue
        for r, c in re.findall(r"\$\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)", expr):
            indices.append((int(r), int(c)))
        for r, c in re.findall(r"(?<!\$)\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)", expr):
            indices.append((int(r), int(c)))
        for r in re.findall(r"\$(\d+)", expr):
            indices.append((int(r), 0))
    return indices


def run_quality_checks(
    rules_json: JsonDict,
    details_json: Optional[JsonDict] = None,
) -> JsonDict:
    issues: List[JsonDict] = []
    metrics = {
        "rules_checked": 0,
        "equation_sides_null": 0,
        "placeholder_tokens": 0,
        "index_out_of_range": 0,
    }

    window_shapes: Dict[str, Tuple[int, int]] = {}
    for w in _iter_window_results(details_json):
        info = w.get("window_info") or {}
        win_id = info.get("window_id")
        shape = _get_window_shape(w)
        if win_id is not None and shape:
            window_shapes[str(win_id)] = shape

    for key, group in (rules_json or {}).items():
        if not isinstance(group, dict):
            continue
        passed_rules = group.get("passed_rules", [])
        if isinstance(passed_rules, dict):
            passed_rules = list(passed_rules.values())
        if not isinstance(passed_rules, list):
            continue

        for rule in passed_rules:
            if not isinstance(rule, dict):
                continue
            metrics["rules_checked"] += 1
            eq = rule.get("equation") or rule.get("rule") or rule.get("description") or ""
            eq_norm = normalize_equation(str(eq))

            if check_equation_sides_null(rule):
                metrics["equation_sides_null"] += 1
                issues.append({
                    "type": "equation_sides_null",
                    "start_loc": str(key),
                    "row_name": group.get("start_loc_row_name"),
                    "equation": eq,
                })

            if check_placeholder_tokens(eq):
                metrics["placeholder_tokens"] += 1
                issues.append({
                    "type": "placeholder_tokens",
                    "start_loc": str(key),
                    "row_name": group.get("start_loc_row_name"),
                    "equation": eq,
                })

            # Index range checks if details_json provided
            if window_shapes:
                sides = rule.get("equation_sides")
                if isinstance(sides, list) and len(sides) == 2:
                    exprs = [str(sides[0]), str(sides[1])]
                else:
                    split = try_split_equation_sides(eq)
                    exprs = [split[0], split[1]] if split else []
                if exprs:
                    indices = _extract_indices(exprs)
                    for sw in rule.get("supporting_windows", []) or []:
                        win_id = sw.get("window_id")
                        if win_id is None:
                            continue
                        shape = window_shapes.get(str(win_id))
                        if not shape:
                            continue
                        rows, cols = shape
                        for r, c in indices:
                            if r < 0 or c < 0 or r >= rows or c >= cols:
                                metrics["index_out_of_range"] += 1
                                issues.append({
                                    "type": "index_out_of_range",
                                    "window_id": win_id,
                                    "start_loc": str(key),
                                    "row_name": group.get("start_loc_row_name"),
                                    "equation": eq_norm,
                                    "index": (r, c),
                                    "window_shape": shape,
                                })

    return {
        "summary": {
            "issues": len(issues),
        },
        "metrics": metrics,
        "issues": issues,
    }


def check_equation_sides_null(rule_entry: JsonDict) -> bool:
    sides = rule_entry.get("equation_sides")
    if sides is None:
        return True
    if not isinstance(sides, list):
        return True
    if len(sides) != 2:
        return True
    if not all(isinstance(x, str) and x.strip() for x in sides):
        return True
    return False


def check_placeholder_tokens(equation: str) -> bool:
    if not isinstance(equation, str):
        return False
    tokens = ["左邊", "右邊", "left", "right"]
    eq = equation.lower()
    for t in tokens:
        if t.lower() in eq:
            return True
    return False
