# -*- coding: utf-8 -*-
"""
Suspect cell aggregation from failed rules.
"""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple
import re

from rule_audit_types import JsonDict, PassedRuleDict
from rule_audit_normalize import try_split_equation_sides


def _col_to_excel(col_idx: int) -> str:
    if col_idx < 26:
        return chr(ord("A") + col_idx)
    first_char = chr(ord("A") + (col_idx // 26) - 1)
    second_char = chr(ord("A") + (col_idx % 26))
    return first_char + second_char


def _to_excel_cell(row_idx: int, col_idx: int) -> str:
    return f"{_col_to_excel(col_idx)}{row_idx + 1}"


def _extract_indices(exprs: List[str]) -> List[Tuple[int, int]]:
    indices: List[Tuple[int, int]] = []
    for expr in exprs:
        if not isinstance(expr, str):
            continue
        # $(r,c)
        for r, c in re.findall(r"\$\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)", expr):
            indices.append((int(r), int(c)))
        # bare (r,c)
        for r, c in re.findall(r"(?<!\$)\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)", expr):
            indices.append((int(r), int(c)))
        # legacy $i -> (i,0)
        for r in re.findall(r"\$(\d+)", expr):
            indices.append((int(r), 0))
    return indices


def _get_equation_sides_from_rule(rule: JsonDict) -> Optional[List[str]]:
    sides = rule.get("equation_sides")
    if isinstance(sides, list) and len(sides) == 2:
        return [str(sides[0]), str(sides[1])]
    eq = rule.get("equation") or rule.get("rule") or rule.get("description")
    if isinstance(eq, str):
        split = try_split_equation_sides(eq)
        if split:
            return [split[0], split[1]]
    return None


def collect_suspect_cells(
    failed_items: List[JsonDict],
    *,
    top_k: int = 50,
    rule_source: Optional[List[PassedRuleDict]] = None,
) -> List[JsonDict]:
    suspect_map: Dict[str, JsonDict] = {}

    for item in failed_items:
        position = item.get("position") or {}
        start_row = position.get("start_row")
        start_col = position.get("start_col")
        window_id = item.get("window_id")
        failed_rules = item.get("failed_rules", [])

        for fr in failed_rules:
            sides = _get_equation_sides_from_rule(fr)
            if not sides:
                continue
            indices = _extract_indices(sides)
            for r, c in indices:
                abs_row = None
                abs_col = None
                excel_cell = None
                if start_row is not None and start_col is not None:
                    abs_row = int(start_row) + r
                    abs_col = int(start_col) + c
                    if abs_row >= 0 and abs_col >= 0:
                        excel_cell = _to_excel_cell(abs_row, abs_col)
                key = f"{abs_row},{abs_col}" if abs_row is not None and abs_col is not None else f"{r},{c}"
                entry = suspect_map.get(key)
                if not entry:
                    entry = {
                        "row_offset": r,
                        "col_offset": c,
                        "abs_row": abs_row,
                        "abs_col": abs_col,
                        "excel_cell": excel_cell,
                        "count": 0,
                        "windows": [],
                        "rules": [],
                    }
                    suspect_map[key] = entry
                entry["count"] += 1
                if window_id is not None and window_id not in entry["windows"]:
                    entry["windows"].append(window_id)
                eq = fr.get("equation") or fr.get("rule")
                if eq and eq not in entry["rules"]:
                    entry["rules"].append(eq)

    ranked = rank_suspect_cells(list(suspect_map.values()))
    return ranked[:top_k] if top_k and top_k > 0 else ranked


def rank_suspect_cells(cells: List[JsonDict]) -> List[JsonDict]:
    return sorted(
        cells,
        key=lambda c: (c.get("count", 0), len(c.get("windows", []))),
        reverse=True,
    )
