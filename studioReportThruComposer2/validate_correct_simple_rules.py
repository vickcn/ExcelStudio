#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
依 correct_simple_hinted_rules.json 驗證 correct_simple.xlsx 之數值關係。
相關註解標明各檢查對應之業務/幾何意義（與 JSON 內 description 一致）。
"""

import json
import math
import sys
from pathlib import Path

import pandas as pd

_REPORT_DIR = Path(__file__).resolve().parent
_RULES_PATH = _REPORT_DIR / "correct_simple_hinted_rules.json"


def _num(x):
    """將儲存格轉成可計算之 float；無法轉換則拋錯以便發現髒資料。"""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        raise ValueError("empty cell")
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(",", "")
    return float(s)


def _close(a, b, tol):
    return abs(a - b) <= tol


def _cell(df, row, col):
    return df.iat[row, col]


def main():
    # 驗證目標：載入規則定義與工作表矩陣（header=None 與規則列索引一致）
    with open(_RULES_PATH, "r", encoding="utf-8") as f:
        spec = json.load(f)
    excel_path = spec["source_excel"]
    tol = float(spec.get("numeric_tolerance", 0.02))
    data_cols = spec["data_column_indices_zero_based"]
    rules = spec["rules"]

    df = pd.read_excel(excel_path, sheet_name=spec.get("sheet", 0), header=None)

    failures = []
    passed = 0

    for rule in rules:
        rid = rule["id"]
        op = rule["op"]

        skip_cols = set(rule.get("skip_columns_zero_based") or [])
        for c in data_cols:
            if c in skip_cols:
                continue
            try:
                bad = False
                if op == "eq":
                    # 驗證目標：兩列在同一資料欄應相等（如進表重量銜接）
                    a = _num(_cell(df, rule["left_row"], c))
                    b = _num(_cell(df, rule["right_row"], c))
                    if not _close(a, b, tol):
                        failures.append((rid, c, f"eq: {a} vs {b}"))
                        bad = True

                elif op == "sub_eq":
                    # 驗證目標：減法平衡（後重 - 零件/膠量 = 下一階後重）
                    m = _num(_cell(df, rule["minuend_row"], c))
                    s = _num(_cell(df, rule["subtrahend_row"], c))
                    r = _num(_cell(df, rule["result_row"], c))
                    if not _close(m - s, r, tol):
                        failures.append((rid, c, f"sub: {m}-{s}!={r}"))
                        bad = True

                elif op == "add_eq":
                    # 驗證目標：加法平衡（前階 + 加工增量 = 後階）
                    a = _num(_cell(df, rule["a_row"], c))
                    b = _num(_cell(df, rule["b_row"], c))
                    r = _num(_cell(df, rule["result_row"], c))
                    if not _close(a + b, r, tol):
                        failures.append((rid, c, f"add: {a}+{b}!={r}"))
                        bad = True

                elif op == "all_eq":
                    # 驗證目標：雕刻線溝段在增量皆為 0 時，多列「後重」應同值
                    vals = [_num(_cell(df, r, c)) for r in rule["rows"]]
                    ref = vals[0]
                    if any(not _close(v, ref, tol) for v in vals[1:]):
                        failures.append((rid, c, f"all_eq: {vals}"))
                        bad = True

                elif op == "eq_when_second_zero":
                    # 驗證目標：壓面量為 0 時，壓面後重量應等於 AC 後段入貨重
                    gate = _num(_cell(df, rule["gate_row"], c))
                    if _close(gate, 0.0, tol):
                        a = _num(_cell(df, rule["left_row"], c))
                        b = _num(_cell(df, rule["right_row"], c))
                        if not _close(a, b, tol):
                            failures.append((rid, c, f"eq_when_zero: {a} vs {b}"))
                            bad = True

                else:
                    failures.append((rid, c, f"unknown op {op}"))
                    bad = True

                if not bad:
                    passed += 1
            except Exception as ex:
                failures.append((rid, c, str(ex)))

    print("rules_file:", _RULES_PATH)
    print("excel:", excel_path)
    print("checks_ok:", passed)
    print("failures:", len(failures))
    for item in failures:
        print(" FAIL", item)
    if failures:
        sys.exit(1)
    sys.exit(0)


if __name__ == "__main__":
    main()
