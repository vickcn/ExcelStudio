from __future__ import annotations

from typing import Any, Dict, List


def build_discovered_rules(inspect_result: Dict[str, Any], *, min_relation_confidence: float = 0.85) -> Dict[str, Any]:
    rules: List[Dict[str, Any]] = []
    rid = 1
    for sheet in inspect_result.get("sheets", []):
        for table in sheet.get("tables", []):
            table_id = table["table_id"]
            rules.append({"rule_id": f"R{rid:04d}", "type": "table_region", "sheet": sheet["sheet"], "table_id": table_id, "range": table["range"]})
            rid += 1
        for layout in sheet.get("layouts", []):
            table_id = layout["table_id"]
            if layout.get("header", {}).get("rows"):
                rules.append({"rule_id": f"R{rid:04d}", "type": "header", "sheet": sheet["sheet"], "table_id": table_id, "header_rows": layout["header"]["rows"]})
                rid += 1
            if layout.get("index", {}).get("cols"):
                rules.append({"rule_id": f"R{rid:04d}", "type": "index", "sheet": sheet["sheet"], "table_id": table_id, "index_cols": layout["index"]["cols"]})
                rid += 1
        for prof in sheet.get("type_profiles", []):
            for col in prof.get("columns", []):
                if col.get("type_confidence", 0) >= 0.8 and col.get("inferred_type") not in {"empty", "mixed"}:
                    rules.append({"rule_id": f"R{rid:04d}", "type": "data_type", "sheet": sheet["sheet"], "table_id": prof["table_id"], "col": col["col"], "label": col["label"], "expected_type": col["inferred_type"]})
                    rid += 1
        for rel in sheet.get("relations", []):
            if rel.get("confidence", 0) >= min_relation_confidence:
                rules.append({"rule_id": f"R{rid:04d}", "type": "arithmetic_relation", "sheet": sheet["sheet"], "axis": rel["axis"], "formula": rel["formula_pattern"], "lag_k": rel.get("lag_k"), "support_rate": rel.get("support_rate"), "tolerance": {"abs": 1e-6, "rel": 1e-4}})
                rid += 1
    return {"version": "excel_audit_rule_v1", "workbook": inspect_result.get("workbook", {}).get("filename"), "rules": rules}
