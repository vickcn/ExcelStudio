from __future__ import annotations

from typing import Any, Dict, List

import pandas as pd

from .cell_types import aggregate_type, infer_cell_type
from .utils import clean_text, excel_cell, is_empty_value


def profile_data_region(sheet_df: pd.DataFrame, table: Dict[str, Any], layout: Dict[str, Any], *, max_anomalies: int = 50) -> Dict[str, Any]:
    region = layout["data_region"]
    if not region.get("range"):
        return {"table_id": table["table_id"], "sheet": table["sheet"], "columns": [], "rows": [], "anomalies": []}

    top = region["top_row"] - 1
    left = region["left_col"] - 1
    bottom = region["bottom_row"] - 1
    right = region["right_col"] - 1
    data = sheet_df.iloc[top : bottom + 1, left : right + 1]
    labels = {x["relative_col"]: x["label"] for x in layout.get("column_labels", [])}
    table_left = table["left_col"] - 1

    col_profiles: List[Dict[str, Any]] = []
    anomalies: List[Dict[str, Any]] = []
    for offset, c in enumerate(range(left, right + 1)):
        values = list(data.iloc[:, offset])
        types = [infer_cell_type(v) for v in values]
        inferred, conf = aggregate_type(types)
        missing = sum(1 for v in values if is_empty_value(v))
        non_empty = len(values) - missing
        type_counts = {t: types.count(t) for t in sorted(set(types))}
        label = labels.get(c - table_left, f"col_{c + 1}")
        examples = [clean_text(v) for v in values if not is_empty_value(v)][:3]
        col_profiles.append(
            {
                "col": c + 1,
                "label": label,
                "inferred_type": inferred,
                "type_confidence": round(conf, 4),
                "non_empty_count": non_empty,
                "missing_count": missing,
                "missing_rate": round(missing / max(1, len(values)), 4),
                "type_counts": type_counts,
                "examples": examples,
            }
        )
        for r_offset, (v, t) in enumerate(zip(values, types)):
            if len(anomalies) >= max_anomalies:
                break
            if inferred != "mixed" and t != "empty":
                bad = False
                if inferred == "number" and t not in {"integer", "number", "percentage"}:
                    bad = True
                elif inferred == "datetime" and t not in {"date", "datetime", "time"}:
                    bad = True
                elif inferred not in {"number", "datetime"} and t != inferred and conf >= 0.85:
                    bad = True
                if bad:
                    anomalies.append(
                        {
                            "cell": excel_cell(top + r_offset + 1, c + 1),
                            "value": clean_text(v),
                            "expected_type": inferred,
                            "actual_type": t,
                            "reason": f"expected {inferred} but got {t}",
                        }
                    )
    row_profiles = []
    for offset, r in enumerate(range(top, bottom + 1)):
        values = list(data.iloc[offset, :])
        types = [infer_cell_type(v) for v in values]
        inferred, conf = aggregate_type(types)
        missing = sum(1 for v in values if is_empty_value(v))
        row_profiles.append(
            {
                "row": r + 1,
                "inferred_type": inferred,
                "type_confidence": round(conf, 4),
                "missing_count": missing,
                "missing_rate": round(missing / max(1, len(values)), 4),
            }
        )
    return {"table_id": table["table_id"], "sheet": table["sheet"], "columns": col_profiles, "rows": row_profiles, "anomalies": anomalies}
