# -*- coding: utf-8 -*-
"""
規律審查工具：共用型別別名與常數（骨架階段，實作時再收斂結構）。

對應開發計畫：tmp2/開發計畫.md
"""

from typing import Any, Dict, List, Optional, Union

# JSON 載入後常見頂層結構（discovered_rules 為 start_loc -> 群組；observed 可能為路徑鍵映射）
JsonValue = Union[Dict[str, Any], List[Any], str, int, float, bool, None]
JsonDict = Dict[str, Any]

# discovered_rules.json 內 passed_rules 單筆（欄位以實際輸出為準，此處僅註記語意）
PassedRuleDict = JsonDict

# analysis_details 內單一視窗結果（與 run_math_rule_analysis 輸出對齊，細節由實作確認）
WindowResultDict = JsonDict

# 品質檢查、比較報告等統一回傳骨架（6.1 機器可讀 JSON）
AuditReportDict = JsonDict
