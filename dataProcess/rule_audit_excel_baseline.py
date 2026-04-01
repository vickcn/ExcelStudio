# -*- coding: utf-8 -*-
"""
Use baseline discovered_rules.json to audit an Excel file and find suspect cells.
No LLM calls are used.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from universal_table_detector import UniversalTableDetector
from window_scanner import WindowScanner

import rule_audit_io
import rule_audit_rules
import rule_audit_eval
import rule_audit_suspect


def _load_main_dataframe(excel_path: Path) -> Any:
    detector = UniversalTableDetector()
    table_results = detector.detect_tables_by_analysis(
        str(excel_path),
        table_mode="pure_numeric",
        use_llm=False,
    )
    if not table_results:
        raise RuntimeError("no table detected")

    dataframes = detector.extract_dataframes(str(excel_path), table_results)
    if not dataframes:
        raise RuntimeError("no dataframe extracted")

    return dataframes[0]


def _build_window_details(windows: List[Dict[str, Any]], scanner: WindowScanner) -> Tuple[List[Dict[str, Any]], int]:
    details: List[Dict[str, Any]] = []
    skipped = 0
    for w in windows:
        prompt = scanner.generate_prompt_data(w)
        if not prompt.get("has_numeric"):
            skipped += 1
            continue
        details.append({
            "window_info": {
                "window_id": w.get("window_id"),
                "start_loc_row_name": w.get("start_loc_row_name"),
                "excel_range": w.get("excel_range"),
                "position": w.get("position"),
            },
            "prompt_data": {
                "values": prompt.get("values"),
            },
        })
    return details, skipped


def _prune_failed_windows(
    failed_windows: List[Dict[str, Any]],
    *,
    include_failed_rules: bool,
    include_rule_meta: bool,
) -> List[Dict[str, Any]]:
    if not include_failed_rules:
        return [
            {
                "window_id": w.get("window_id"),
                "start_loc_row_name": w.get("start_loc_row_name"),
                "excel_range": w.get("excel_range"),
                "position": w.get("position"),
            }
            for w in failed_windows
        ]

    pruned = []
    for w in failed_windows:
        entry = {
            "window_id": w.get("window_id"),
            "start_loc_row_name": w.get("start_loc_row_name"),
            "excel_range": w.get("excel_range"),
            "position": w.get("position"),
            "failed_rules": [],
        }
        for fr in w.get("failed_rules", []):
            fr_item = {
                "equation": fr.get("equation"),
                "equation_sides": fr.get("equation_sides"),
                "left_value": fr.get("left_value"),
                "right_value": fr.get("right_value"),
                "diff": fr.get("diff"),
                "tolerance": fr.get("tolerance"),
                "error": fr.get("error"),
            }
            if include_rule_meta:
                fr_item["rule_meta"] = fr.get("rule_meta")
            entry["failed_rules"].append(fr_item)
        pruned.append(entry)
    return pruned


def _collect_row_stats(
    details: List[Dict[str, Any]],
    baseline_rules: List[Dict[str, Any]],
    *,
    tolerance: float,
    strict_row_match: bool,
) -> Dict[str, Dict[str, Any]]:
    row_stats: Dict[str, Dict[str, Any]] = {}
    for w in details:
        info = w.get("window_info") or {}
        row_name = info.get("start_loc_row_name")
        if row_name is None:
            row_name = "unknown"

        rules_to_use = [
            r for r in baseline_rules
            if r.get("_row_name") in (None, "", row_name)
        ]
        if not rules_to_use:
            if strict_row_match:
                continue
            rules_to_use = baseline_rules

        values = (w.get("prompt_data") or {}).get("values") or []

        stats = row_stats.setdefault(row_name, {
            "windows": 0,
            "rule_checks": 0,
            "passed_checks": 0,
            "failed_checks": 0,
        })
        stats["windows"] += 1

        for rule in rules_to_use:
            stats["rule_checks"] += 1
            eval_result = rule_audit_eval.evaluate_rule_on_window(
                rule,
                values,
                tolerance=tolerance,
            )
            if eval_result.get("ok"):
                stats["passed_checks"] += 1
            else:
                stats["failed_checks"] += 1

    # compute pass rate
    for row_name, stats in row_stats.items():
        total = stats["rule_checks"]
        stats["pass_rate"] = stats["passed_checks"] / total if total else 0.0
    return row_stats


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Audit Excel by baseline discovered_rules.json (no LLM)."
    )
    parser.add_argument("--baseline-rules", required=True, type=Path, help="A 的 discovered_rules.json")
    parser.add_argument("--excel", required=True, type=Path, help="要檢查的 Excel 檔")
    parser.add_argument("--row-name", default=None, help="只檢查指定 start_loc_row_name")
    parser.add_argument("--window-height", type=int, default=3)
    parser.add_argument("--window-width", type=int, default=1)
    parser.add_argument("--tolerance", type=float, default=0.01)
    parser.add_argument("--strict-row-match", action="store_true", help="只套用相同 row 的規則")
    parser.add_argument(
        "--output-mode",
        choices=("minimal", "windows", "full"),
        default="minimal",
        help="輸出精簡模式：minimal(只含suspect_cells)、windows(含失敗視窗)、full(含rule_meta)",
    )
    parser.add_argument("--top-k", type=int, default=50, help="只輸出前 K 筆可疑儲存格（0 表示全部）")
    parser.add_argument("--out", required=True, type=Path, help="輸出 JSON")

    args = parser.parse_args()

    excel_path = args.excel
    if not excel_path.is_absolute():
        excel_path = Path.cwd() / excel_path
    if not excel_path.exists():
        raise SystemExit(f"excel not found: {excel_path}")

    baseline_json = rule_audit_io.load_json(args.baseline_rules)
    baseline_rules = rule_audit_rules.extract_passed_rules(baseline_json)
    if not baseline_rules:
        raise SystemExit("baseline rules empty")

    main_df = _load_main_dataframe(excel_path)
    scanner = WindowScanner((args.window_height, args.window_width))
    windows = scanner.scan_dataframe(main_df)
    if args.row_name:
        windows = scanner.filter_windows_by_start_loc_row(windows, args.row_name)

    details, skipped = _build_window_details(windows, scanner)

    audit_report = rule_audit_eval.audit_windows_with_baseline(
        details,
        baseline_rules,
        row_name=args.row_name,
        tolerance=args.tolerance,
        strict_row_match=args.strict_row_match,
    )

    suspect_cells = rule_audit_suspect.collect_suspect_cells(
        audit_report.get("failed_windows", []),
        top_k=args.top_k,
    )

    # row mapping (row_name -> baseline id)
    row_to_id = {}
    for key, group in (baseline_json or {}).items():
        if not isinstance(group, dict):
            continue
        rn = group.get("start_loc_row_name")
        if rn and rn not in row_to_id:
            row_to_id[rn] = str(key)

    row_stats = _collect_row_stats(
        details,
        baseline_rules,
        tolerance=args.tolerance,
        strict_row_match=args.strict_row_match,
    )

    # print summary per row
    print(f"\n=== Row Rule Pass Summary ===")
    print(f"file: {excel_path}")
    for row_name, stats in sorted(row_stats.items(), key=lambda x: x[0]):
        baseline_id = row_to_id.get(row_name, "n/a")
        print(
            f"- row_name={row_name} baseline_id={baseline_id} "
            f"windows={stats['windows']} "
            f"rule_checks={stats['rule_checks']} "
            f"pass_rate={stats['pass_rate']:.2%}"
        )

    report = {
        "meta": {
            "excel_file": str(excel_path),
            "baseline_rules": str(args.baseline_rules),
            "window_shape": [args.window_height, args.window_width],
            "row_name": args.row_name,
            "tolerance": args.tolerance,
            "strict_row_match": args.strict_row_match,
            "total_windows": len(windows),
            "skipped_windows_no_numeric": skipped,
            "evaluated_windows": len(details),
        },
        "summary": audit_report.get("summary", {}),
        "metrics": audit_report.get("metrics", {}),
        "row_rule_pass_summary": {
            row_name: {
                "baseline_id": row_to_id.get(row_name),
                "windows": stats["windows"],
                "rule_checks": stats["rule_checks"],
                "pass_rate": stats["pass_rate"],
            }
            for row_name, stats in row_stats.items()
        },
        "suspect_cells": suspect_cells,
    }

    if args.output_mode in ("windows", "full"):
        report["failed_windows"] = _prune_failed_windows(
            audit_report.get("failed_windows", []),
            include_failed_rules=True,
            include_rule_meta=(args.output_mode == "full"),
        )

    rule_audit_io.write_report(report, args.out)


if __name__ == "__main__":
    main()
