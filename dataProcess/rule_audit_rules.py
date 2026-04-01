# -*- coding: utf-8 -*-
"""
規律審查工具：從 discovered_rules 抽取規律、集合比較。

職責：支援 compare-rules 子命令；可選 row 群組（start_loc）與一致性門檻過濾。
"""

from typing import Any, Dict, List, Optional

from rule_audit_types import JsonDict, PassedRuleDict
from rule_audit_normalize import normalize_equation


def extract_passed_rules(
    discovered_json: JsonDict,
    row_key: Optional[str] = None,
    min_consistency: Optional[float] = None,
) -> List[PassedRuleDict]:
    """
    從 discovered_rules.json 結構抽出「通過規律」列表。

    Args:
        discovered_json: load_json(discovered_rules.json) 結果。
        row_key: 若指定，只取該 start_loc（或計畫中約定之群組鍵）底下之 passed_rules。
        min_consistency: 若指定，過濾 consistency_rate 未達門檻之規律。

    Returns:
        passed_rules 扁平化列表（或依實作約定附帶群組欄位）。
    """
    if not isinstance(discovered_json, dict):
        return []
    results: List[PassedRuleDict] = []
    for key, group in discovered_json.items():
        if not isinstance(group, dict):
            continue
        if row_key is not None:
            key_str = str(key)
            row_name = str(group.get("start_loc_row_name", ""))
            if str(row_key) not in (key_str, row_name):
                continue

        passed_rules = group.get("passed_rules", [])
        if isinstance(passed_rules, dict):
            passed_rules = list(passed_rules.values())
        if not isinstance(passed_rules, list):
            continue

        for rule in passed_rules:
            if not isinstance(rule, dict):
                continue
            if min_consistency is not None:
                try:
                    if float(rule.get("consistency_rate", 0.0)) < float(min_consistency):
                        continue
                except Exception:
                    continue
            enriched = dict(rule)
            enriched["_row_key"] = str(key)
            enriched["_row_name"] = group.get("start_loc_row_name")
            enriched["_window_count"] = group.get("window_count")
            results.append(enriched)
    return results


def compare_rule_sets(
    a_rules: List[PassedRuleDict],
    b_rules: List[PassedRuleDict],
    *,
    equation_key: str = "equation",
) -> JsonDict:
    """
    比較 A/B 兩組已抽出之規律集合。

    Args:
        a_rules: baseline 或 A 檔規律列表。
        b_rules: B 檔規律列表。
        equation_key: 比對主鍵欄位名稱（預設 equation；可改 rule）。

    Returns:
        建議含 summary、metrics、shared_rules、a_only_rules、b_only_rules（開發計畫 6.1）。
    """
    def _bucket(rules: List[PassedRuleDict]) -> Dict[str, List[PassedRuleDict]]:
        buckets: Dict[str, List[PassedRuleDict]] = {}
        for r in rules:
            key = rule_identity(r, equation_key=equation_key)
            if not key:
                continue
            buckets.setdefault(key, []).append(r)
        return buckets

    a_map = _bucket(a_rules)
    b_map = _bucket(b_rules)

    a_keys = set(a_map.keys())
    b_keys = set(b_map.keys())
    shared_keys = sorted(a_keys & b_keys)
    a_only_keys = sorted(a_keys - b_keys)
    b_only_keys = sorted(b_keys - a_keys)

    def _pack(key: str, rules: List[PassedRuleDict], label: str) -> Dict[str, Any]:
        sample = rules[0] if rules else {}
        return {
            "equation": key,
            f"{label}_count": len(rules),
            f"{label}_sample": sample,
        }

    shared_rules = []
    for key in shared_keys:
        shared_rules.append({
            "equation": key,
            "a_count": len(a_map.get(key, [])),
            "b_count": len(b_map.get(key, [])),
            "a_sample": (a_map.get(key) or [{}])[0],
            "b_sample": (b_map.get(key) or [{}])[0],
        })

    a_only_rules = [_pack(k, a_map[k], "a") for k in a_only_keys]
    b_only_rules = [_pack(k, b_map[k], "b") for k in b_only_keys]

    metrics = {
        "a_rules": len(a_rules),
        "b_rules": len(b_rules),
        "shared_rules": len(shared_rules),
        "a_only_rules": len(a_only_rules),
        "b_only_rules": len(b_only_rules),
    }
    summary = {
        "shared_ratio_vs_a": metrics["shared_rules"] / metrics["a_rules"] if metrics["a_rules"] else 0.0,
        "shared_ratio_vs_b": metrics["shared_rules"] / metrics["b_rules"] if metrics["b_rules"] else 0.0,
    }

    return {
        "summary": summary,
        "metrics": metrics,
        "shared_rules": shared_rules,
        "a_only_rules": a_only_rules,
        "b_only_rules": b_only_rules,
    }


def rule_identity(rule: PassedRuleDict, equation_key: str = "equation") -> Any:
    """
    由單筆規律產生「集合比對用」鍵（可呼叫 normalize_equation）。

    實作時與 compare_rule_sets 共用，避免重複邏輯。
    """
    eq = None
    if isinstance(rule, dict):
        eq = rule.get(equation_key)
        if eq is None:
            eq = rule.get("rule") or rule.get("description")
    return normalize_equation(str(eq)) if eq is not None else ""
