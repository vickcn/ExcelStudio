# -*- coding: utf-8 -*-
"""
規律審查工具：equation / 規律字串正規化。

職責：統一空白、括號、索引表示（$0 / $(0,0) 等）以利集合比對；對策見開發計畫 9. 風險。
"""

from typing import Optional, Tuple
import re


def normalize_equation(eq: str) -> str:
    """
    將 equation 或規律字串正規化為可比對形式。

    Args:
        eq: 原始字串（可能含空白差異、不同索引寫法）。

    Returns:
        正規化後字串；實作時需與 extract_passed_rules / compare_rule_sets 約定一致。

    Note:
        開發計畫建議：規律文字重複但空白格式不同時，統一由此函式處理。
    """
    if not isinstance(eq, str):
        return ""
    s = eq.strip()
    if not s:
        return ""
    # Normalize full-width equals and remove code fences if present.
    s = s.replace("＝", "=")
    s = s.replace("```", " ").replace("`", " ")
    # Normalize "$ (" -> "$("
    s = re.sub(r"\$\s+\(", "$(", s)
    # Normalize bare "(r,c)" into "$(r,c)"
    s = re.sub(r"(?<!\$)\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)", r"$(\1,\2)", s)
    # Normalize legacy "$i" into "$(i,0)"
    s = re.sub(r"\$(\d+)", r"$(\1,0)", s)
    # Collapse whitespace
    s = re.sub(r"\s+", " ", s).strip()
    # For identity comparison, remove all spaces
    s = re.sub(r"\s+", "", s)
    return s


def try_split_equation_sides(equation: str) -> Optional[Tuple[str, str]]:
    """
    由單一 equation 字串嘗試拆出左右式（無 '=' 或格式不支援則回傳 None）。

    用途：equation_sides 為 null 時的補救（開發計畫 9.）；品質檢查可沿用。
    """
    if not isinstance(equation, str):
        return None
    s = equation.replace("＝", "=").strip()
    if "=" not in s:
        return None
    left, right = s.split("=", 1)
    left = left.strip()
    right = right.strip()
    if not left or not right:
        return None
    left_norm = normalize_equation(left)
    right_norm = normalize_equation(right)
    if not left_norm or not right_norm:
        return None
    return left_norm, right_norm
