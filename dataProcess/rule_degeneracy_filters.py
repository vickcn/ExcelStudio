# -*- coding: utf-8 -*-
"""
無意義規律過濾。

策略 1：代數／結構退化（恒等式、A+B=B 強制 A=0、A*B+C=C 強制乘積為 0 等）。
策略 2：各 $(r,c) 在跨視窗樣本上須有足夠非零樣本，避免 0 退化假成立。

預設參數與 run_math_rule_analysis 內 analyzer 設定對齊，可直接用於擋常見 trivial 規律。
"""

from typing import Any, Dict, List, Optional, Tuple
import re

from rule_audit_normalize import normalize_equation, try_split_equation_sides

_PLACEHOLDER_RE = re.compile(r"\$\(-?\d+,-?\d+\)")

# 與 should_reject_passed_rule 預設一致；可由 MathRuleAnalyzer 覆寫後傳入。
DEFAULT_MIN_NONZERO_COUNT = 1
DEFAULT_MIN_DISTINCT_NONZERO = 1


def _is_single_placeholder_norm(s: str) -> bool:
    if not isinstance(s, str) or not s:
        return False
    return bool(_PLACEHOLDER_RE.fullmatch(s.strip()))


def _extract_placeholders(expr: str) -> List[str]:
    if not isinstance(expr, str):
        return []
    norm = normalize_equation(expr)
    if not norm:
        return []
    return _PLACEHOLDER_RE.findall(norm)


def _parse_placeholder(token: str) -> Optional[Tuple[int, int]]:
    if not isinstance(token, str):
        return None
    m = re.match(r"^\$\((-?\d+),(-?\d+)\)$", token.strip())
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def _coerce_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        text = str(value).strip().replace(",", "")
        if text == "":
            return None
        return float(text)
    except Exception:
        return None


def equation_is_structurally_degenerate(
    equation: str,
    equation_sides: Optional[List[str]] = None,
) -> Tuple[bool, str]:
    """
    策略 1：僅依方程式字串與 equation_sides 做結構判斷。

    Args:
        equation: 完整等式，例如 "$(0,0)*$(1,0)+$(2,0)=$(2,0)"
        equation_sides: 可選，左右式各一項的列表

    Returns:
        (是否應剔除, 原因代碼或簡短說明)。框架預設 (False, "")。

    實作提示：
        - 可先用 try_split_equation_sides 取得左右式，再 parse 出所有 $(r,c) 占位符。
        - 檢查 LHS 與 RHS 是否代數上等價於「恒等」或「強制某乘積為 0」等形式。
    """
    normalized = normalize_equation(equation)
    if not normalized:
        return False, ""
    if not _PLACEHOLDER_RE.search(normalized):
        return True, "no_operands"

    sides: Optional[Tuple[str, str]] = None
    if isinstance(equation_sides, list) and len(equation_sides) == 2:
        left = normalize_equation(str(equation_sides[0]))
        right = normalize_equation(str(equation_sides[1]))
        if left and right:
            sides = (left, right)
    if sides is None:
        sides = try_split_equation_sides(equation)

    if sides:
        left, right = sides
        if left == right:
            return True, "identity_equation"

        if _is_single_placeholder_norm(right):
            suf = "+" + right
            if left.endswith(suf):
                head = left[: -len(suf)]
                if head and _is_single_placeholder_norm(head):
                    return True, "add_forced_zero_trailing"
            pre = right + "+"
            if left.startswith(pre):
                tail = left[len(pre) :]
                if tail and _is_single_placeholder_norm(tail):
                    return True, "add_forced_zero_leading"
            if left.startswith(right + "-"):
                rest = left[len(right) + 1 :]
                if rest and "*" not in rest and "/" not in rest and _is_single_placeholder_norm(rest):
                    return True, "sub_minuend_eq_rhs_forces_subtrahend_zero"

    return False, ""


def collect_operand_samples_across_windows(
    group_results: List[Dict[str, Any]],
    equation: str,
    equation_sides: Optional[List[str]],
) -> Dict[str, List[Any]]:
    """
    從同一 start_loc_row 的各視窗中，抽取方程式內 $(r,c) 對應的儲存格原始值。

    Args:
        group_results: 與 run_math_rule_analysis._retrospective_validate_start_loc_rows 相同群組
        equation: 規律 equation
        equation_sides: 可選 sides，便於只掃描相關占位符

    Returns:
        鍵為正規化占位符字串（例如 "$(0,0)"），值為跨視窗收集到的原始值列表。
        框架預設回傳空 dict，由呼叫端視為「略過策略 2」或於實作後填入。

    實作提示：
        - 從 window_result 內既有欄位（例如視窗子矩陣、dataframe 切片）依 (r,c) 取値；
        - 與 llm_math_rule_detector / WindowScanner 約定之索引原點需一致。
    """
    exprs: List[str] = []
    if isinstance(equation_sides, list) and len(equation_sides) == 2:
        exprs = [str(equation_sides[0]), str(equation_sides[1])]
    elif isinstance(equation, str):
        exprs = [equation]

    placeholders: List[str] = []
    seen = set()
    for expr in exprs:
        for token in _extract_placeholders(expr):
            if token not in seen:
                placeholders.append(token)
                seen.add(token)

    if not placeholders:
        return {}

    parsed = []
    for token in placeholders:
        rc = _parse_placeholder(token)
        if rc is not None:
            parsed.append((token, rc[0], rc[1]))

    if not parsed:
        return {}

    samples: Dict[str, List[Any]] = {token: [] for token, _, _ in parsed}

    for window_result in group_results:
        window_info = window_result.get("window_info") or {}
        values = window_info.get("values")
        if not isinstance(values, list):
            continue
        for token, r, c in parsed:
            if r < 0 or c < 0:
                continue
            if r >= len(values):
                continue
            row_vals = values[r]
            if not isinstance(row_vals, list) or c >= len(row_vals):
                continue
            v = _coerce_float(row_vals[c])
            if v is None:
                continue
            samples[token].append(v)

    return samples


def operand_support_is_degenerate(
    samples_by_placeholder: Dict[str, List[Any]],
    min_nonzero_count: int = DEFAULT_MIN_NONZERO_COUNT,
    min_distinct_nonzero: int = DEFAULT_MIN_DISTINCT_NONZERO,
) -> Tuple[bool, str]:
    """
    策略 2：依占位符樣本統計量判斷是否「無訊息」導致假成立。

    Args:
        samples_by_placeholder: collect_operand_samples_across_windows 的輸出
        min_nonzero_count: 至少要有幾筆可判為非零的樣本（實作時定義「非零」）
        min_distinct_nonzero: 非零相異值至少幾個（可區分常數 0 退化）

    Returns:
        (是否應剔除, 原因)。框架預設 (False, "").

    實作提示：
        - 僅對「應作為係數／加減量」的占位符套用較嚴格門檻，需與業務約定或 row 標籤結合時再擴充。
    """
    if not samples_by_placeholder:
        return False, ""

    eps = 1e-12
    for token, values in samples_by_placeholder.items():
        nums = []
        for v in values:
            fv = _coerce_float(v)
            if fv is not None:
                nums.append(fv)
        if not nums:
            return True, f"operand_no_numeric:{token}"
        nonzero = [v for v in nums if abs(v) > eps]
        if len(nonzero) < min_nonzero_count:
            return True, f"operand_nonzero_insufficient:{token}"
        distinct_nonzero = len({round(v, 12) for v in nonzero})
        if distinct_nonzero < min_distinct_nonzero:
            return True, f"operand_distinct_insufficient:{token}"

    return False, ""


def should_reject_passed_rule(
    equation: str,
    equation_sides: Optional[List[str]],
    group_results: List[Dict[str, Any]],
    min_nonzero_count: int = DEFAULT_MIN_NONZERO_COUNT,
    min_distinct_nonzero: int = DEFAULT_MIN_DISTINCT_NONZERO,
) -> Tuple[bool, str]:
    """
    整合策略 1 + 2：通過一致性門檻後、寫入 passed_rules 前呼叫。

    Returns:
        (True, reason) 表示應自 passed_rules 剔除並可記錄於 degeneracy_reject_reason；
        (False, "") 表示保留。

    實作順序建議：先結構退化，再數值支援；任一命中即 reject。
    """
    degenerate, reason = equation_is_structurally_degenerate(equation, equation_sides)
    if degenerate:
        return True, reason or "structural_degenerate"

    samples = collect_operand_samples_across_windows(group_results, equation, equation_sides)
    if samples:
        degenerate2, reason2 = operand_support_is_degenerate(
            samples,
            min_nonzero_count=min_nonzero_count,
            min_distinct_nonzero=min_distinct_nonzero,
        )
        if degenerate2:
            return True, reason2 or "operand_support_degenerate"

    return False, ""
