# -*- coding: utf-8 -*-
"""
Rule audit CLI tool.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import NoReturn, Optional, Sequence

# Allow `python tmp2/dataProcess/rule_audit_tool.py ...`
_DATA_PROCESS_DIR = Path(__file__).resolve().parent
if str(_DATA_PROCESS_DIR) not in sys.path:
    sys.path.insert(0, str(_DATA_PROCESS_DIR))

import rule_audit_eval
import rule_audit_io
import rule_audit_quality
import rule_audit_rules
import rule_audit_suspect


def _die(msg: str, code: int = 2) -> NoReturn:
    print(msg, file=sys.stderr)
    raise SystemExit(code)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Rule audit tool: compare rules, audit baseline, find suspect cells, quality checks.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_cmp = sub.add_parser("compare-rules", help="Compare A/B discovered_rules.json")
    p_cmp.add_argument("--a-rules", required=True, type=Path)
    p_cmp.add_argument("--b-rules", required=True, type=Path)
    p_cmp.add_argument("--row-key", default=None)
    p_cmp.add_argument("--min-consistency", type=float, default=None)
    p_cmp.add_argument("--out", required=True, type=Path)

    p_aud = sub.add_parser("audit-by-baseline", help="Audit analysis_details by baseline rules")
    p_aud.add_argument("--baseline-rules", required=True, type=Path)
    p_aud.add_argument("--target-details", required=True, type=Path)
    p_aud.add_argument("--row-name", default=None)
    p_aud.add_argument("--tolerance", type=float, default=0.01)
    p_aud.add_argument("--strict-row-match", action="store_true")
    p_aud.add_argument("--out", required=True, type=Path)

    p_sus = sub.add_parser("find-suspect-cells", help="Aggregate suspect cells from failed rules")
    p_sus.add_argument("--target-details", required=True, type=Path)
    p_sus.add_argument("--rule-source", default=None, type=Path)
    p_sus.add_argument("--top-k", type=int, default=50)
    p_sus.add_argument("--strict-row-match", action="store_true")
    p_sus.add_argument("--out", required=True, type=Path)

    p_cell = sub.add_parser(
        "audit-by-baseline-cells",
        help="Audit by baseline and output suspect cells",
    )
    p_cell.add_argument("--baseline-rules", required=True, type=Path)
    p_cell.add_argument("--target-details", required=True, type=Path)
    p_cell.add_argument("--row-name", default=None)
    p_cell.add_argument("--tolerance", type=float, default=0.01)
    p_cell.add_argument("--strict-row-match", action="store_true")
    p_cell.add_argument("--out", required=True, type=Path)

    p_q = sub.add_parser("quality-check", help="Quality checks for rule outputs")
    p_q.add_argument("--rules", required=True, type=Path)
    p_q.add_argument("--details", default=None, type=Path)
    p_q.add_argument("--out", required=True, type=Path)

    return parser


def cmd_compare_rules(args: argparse.Namespace) -> None:
    a_data = rule_audit_io.load_json(args.a_rules)
    b_data = rule_audit_io.load_json(args.b_rules)
    a_rules = rule_audit_rules.extract_passed_rules(
        a_data,
        row_key=args.row_key,
        min_consistency=args.min_consistency,
    )
    b_rules = rule_audit_rules.extract_passed_rules(
        b_data,
        row_key=args.row_key,
        min_consistency=args.min_consistency,
    )
    report = rule_audit_rules.compare_rule_sets(a_rules, b_rules)
    rule_audit_io.write_report(report, args.out)


def cmd_audit_by_baseline(args: argparse.Namespace) -> None:
    baseline = rule_audit_io.load_json(args.baseline_rules)
    details = rule_audit_io.load_json(args.target_details)
    rules = rule_audit_rules.extract_passed_rules(baseline)
    report = rule_audit_eval.audit_windows_with_baseline(
        details,
        rules,
        row_name=args.row_name,
        tolerance=args.tolerance,
        strict_row_match=args.strict_row_match,
    )
    rule_audit_io.write_report(report, args.out)


def cmd_find_suspect_cells(args: argparse.Namespace) -> None:
    details = rule_audit_io.load_json(args.target_details)
    rule_src = rule_audit_io.load_optional_json(args.rule_source)
    baseline_list = None
    if rule_src is not None:
        baseline_list = rule_audit_rules.extract_passed_rules(rule_src)

    failed_items = []
    if baseline_list:
        report = rule_audit_eval.audit_windows_with_baseline(
            details,
            baseline_list,
            strict_row_match=args.strict_row_match,
        )
        failed_items = report.get("failed_windows", [])
    else:
        windows = []
        if isinstance(details, list):
            windows = details
        elif isinstance(details, dict):
            windows = details.get("detailed_results") or details.get("window_results") or details.get("results") or []
        for w in windows:
            win_info = w.get("window_info") or {}
            val = w.get("validation") or {}
            validation_details = val.get("validation_details") or []
            failed_rules = []
            for v in validation_details:
                if v.get("is_valid") is True:
                    continue
                eq = v.get("rule") or v.get("equation_preview") or v.get("equation")
                eq_sides = None
                mv = v.get("math_verification") or {}
                if isinstance(mv, dict):
                    eq_sides = mv.get("equation_sides")
                failed_rules.append({
                    "equation": eq,
                    "equation_sides": eq_sides,
                })
            if failed_rules:
                failed_items.append({
                    "window_id": win_info.get("window_id"),
                    "start_loc_row_name": win_info.get("start_loc_row_name"),
                    "excel_range": win_info.get("excel_range"),
                    "position": win_info.get("position"),
                    "failed_rules": failed_rules,
                })

    cells = rule_audit_suspect.collect_suspect_cells(
        failed_items,
        top_k=args.top_k,
        rule_source=baseline_list,
    )
    rule_audit_io.write_report({"suspect_cells": cells, "summary": {}}, args.out)


def cmd_audit_by_baseline_cells(args: argparse.Namespace) -> None:
    baseline = rule_audit_io.load_json(args.baseline_rules)
    details = rule_audit_io.load_json(args.target_details)
    rules = rule_audit_rules.extract_passed_rules(baseline)
    report = rule_audit_eval.audit_details_cells_with_baseline(
        details,
        rules,
        row_name=args.row_name,
        tolerance=args.tolerance,
        strict_row_match=args.strict_row_match,
    )
    rule_audit_io.write_report(report, args.out)


def cmd_quality_check(args: argparse.Namespace) -> None:
    rules = rule_audit_io.load_json(args.rules)
    det = rule_audit_io.load_optional_json(args.details)
    report = rule_audit_quality.run_quality_checks(rules, det)
    rule_audit_io.write_report(report, args.out)


def main(argv: Optional[Sequence[str]] = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if args.command == "compare-rules":
            cmd_compare_rules(args)
        elif args.command == "audit-by-baseline":
            cmd_audit_by_baseline(args)
        elif args.command == "find-suspect-cells":
            cmd_find_suspect_cells(args)
        elif args.command == "audit-by-baseline-cells":
            cmd_audit_by_baseline_cells(args)
        elif args.command == "quality-check":
            cmd_quality_check(args)
        else:
            _die(f"unknown command: {args.command}")
    except NotImplementedError as exc:
        _die(f"not implemented: {exc}")


if __name__ == "__main__":
    main()
