from __future__ import annotations

from typing import Any, Dict, List

from .utils import normalize_label


def infer_semantic_rules(layout: Dict[str, Any], relations: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    labels = [x.get("label", "") for x in layout.get("column_labels", [])]
    norm_map = {normalize_label(x): x for x in labels if x}
    relation_text = {r["formula_pattern"]: r for r in relations if r.get("axis") == "column"}
    candidates: List[Dict[str, Any]] = []

    def has_any(words: list[str]) -> str | None:
        for k, raw in norm_map.items():
            if any(w in k for w in words):
                return raw
        return None

    qty = has_any(["數量", "qty", "quantity"])
    price = has_any(["單價", "價格", "price", "unitprice"])
    amount = has_any(["金額", "amount", "total", "小計"])
    if qty and price and amount:
        candidates.append(_candidate("amount_equals_qty_times_price", "欄名包含數量、單價、金額", f"{amount} = {qty} * {price}", relation_text))

    input_col = has_any(["投入", "input", "進料"])
    good = has_any(["良品", "good", "ok"])
    bad = has_any(["不良", "ng", "defect"])
    if input_col and good and bad:
        candidates.append(_candidate("input_equals_good_plus_bad", "欄名包含投入、良品、不良", f"{input_col} = {good} + {bad}", relation_text))

    rate = has_any(["良率", "yield"])
    if rate and good and input_col:
        candidates.append(_candidate("yield_equals_good_div_input", "欄名包含良率、良品、投入", f"{rate} = {good} / {input_col}", relation_text))

    defect_rate = has_any(["不良率", "defectrate", "ngrate"])
    if defect_rate and bad and input_col:
        candidates.append(_candidate("defect_rate_equals_bad_div_input", "欄名包含不良率、不良、投入", f"{defect_rate} = {bad} / {input_col}", relation_text))

    balance = has_any(["結存", "庫存", "餘額", "balance", "stock"])
    inbound = has_any(["入庫", "收入", "增加", "inbound", "in"])
    outbound = has_any(["出庫", "支出", "減少", "outbound", "out"])
    if balance and inbound and outbound:
        candidates.append({"rule_name": "balance_continuity", "reason": "欄名包含結存/庫存、入庫、出庫", "expected_formula": "本期結存 = 前期結存 + 入庫 - 出庫", "tested": False, "support_rate": None, "status": "candidate_only"})

    return candidates


def _candidate(name: str, reason: str, expected_formula: str, relation_text: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    matched = None
    nf = normalize_label(expected_formula)
    for formula, rel in relation_text.items():
        if normalize_label(formula) == nf:
            matched = rel
            break
    return {
        "rule_name": name,
        "reason": reason,
        "expected_formula": expected_formula,
        "tested": matched is not None,
        "support_rate": matched.get("support_rate") if matched else None,
        "status": "passed" if matched else "not_confirmed",
    }
