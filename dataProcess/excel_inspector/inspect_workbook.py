from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from .header_index_detector import infer_header_index
from .relation_detector import detect_relations
from .rule_exporter import build_discovered_rules
from .rule_inferencer import infer_semantic_rules
from .table_detector import detect_tables_from_sheet
from .type_profiler import profile_data_region
from .utils import safe_json
from .workbook_loader import load_workbook_frames


def inspect_workbook(path: str | Path, *, include_rules: bool = True, relation_max_k: int = 6) -> Dict[str, Any]:
    p = Path(path)
    frames = load_workbook_frames(p)
    result: Dict[str, Any] = {
        "version": "excel_inspector_v1",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "workbook": {"filename": p.name, "path": str(p), "sheet_count": len(frames)},
        "sheets": [],
        "summary": {},
    }
    total_tables = 0
    total_relations = 0
    total_type_anomalies = 0

    for sheet_name, df in frames.items():
        tables = detect_tables_from_sheet(sheet_name, df)
        layouts = []
        profiles = []
        all_relations = []
        semantic_rules = []
        for table in tables:
            layout = infer_header_index(df, table)
            profile = profile_data_region(df, table, layout)
            relations = detect_relations(df, table, layout, max_k=relation_max_k)
            semantics = infer_semantic_rules(layout, relations)
            layouts.append(layout)
            profiles.append(profile)
            all_relations.extend(relations)
            semantic_rules.extend(semantics)
        total_tables += len(tables)
        total_relations += len(all_relations)
        total_type_anomalies += sum(len(pf.get("anomalies", [])) for pf in profiles)
        result["sheets"].append(
            {
                "sheet": sheet_name,
                "shape": {"rows": int(df.shape[0]), "cols": int(df.shape[1])},
                "tables": tables,
                "layouts": layouts,
                "type_profiles": profiles,
                "relations": all_relations,
                "semantic_rule_candidates": semantic_rules,
            }
        )
    result["summary"] = {
        "total_sheets": len(frames),
        "total_tables": total_tables,
        "total_relations": total_relations,
        "total_type_anomalies": total_type_anomalies,
    }
    if include_rules:
        result["discovered_rules"] = build_discovered_rules(result)
    return safe_json(result)


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect complex Excel workbooks and emit JSON.")
    parser.add_argument("xlsx_path", help="Path to .xlsx/.xls/.xlsm/.csv file")
    parser.add_argument("--out", help="Output JSON path. Defaults to stdout.")
    parser.add_argument("--relation-max-k", type=int, default=6)
    args = parser.parse_args()

    result = inspect_workbook(args.xlsx_path, relation_max_k=args.relation_max_k)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.out).write_text(text, encoding="utf-8")
    else:
        print(text)


if __name__ == "__main__":
    main()
