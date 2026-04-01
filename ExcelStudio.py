import argparse
import os
import sys
from typing import Optional, Dict, Any

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PROCESS_DIR = os.path.join(CURRENT_DIR, "dataProcess")
if DATA_PROCESS_DIR not in sys.path:
    sys.path.insert(0, DATA_PROCESS_DIR)

from dataProcess.rule_audit_mark_excel import build_arg_parser as build_arg_parser_rule_audit_mark_excel
from dataProcess.rule_audit_mark_excel import main as main_rule_audit_mark_excel

# 快速標記（Phase1 規則就緒後對 target 產出 audit_marked_fast.xlsx 等）：實作完成後再取消註解並依模組實際路徑調整。
# from dataProcess import rule_audit_mark_fast as rule_audit_mark_fast
# from dataProcess.rule_audit_mark_fast import build_arg_parser as build_arg_parser_rule_audit_mark_fast
# from dataProcess.rule_audit_mark_fast import main as main_rule_audit_mark_fast

from dataProcess import utils as u


def apply_discovered_rules_fast_mark_target(
        target_excel,
        detect_rules_file,
        out_excel_fast=None,
        *,
        baseline_excel=None,
        row_name=None,
        window_height=None,
        window_width=None,
        tolerance=None,
        strict_row_match=False,
        phase1_rules_file=None,
    ):
    """
    Fast-mark target Excel using discovered rules (Phase1 optional).

    This currently reuses rule_audit_mark_excel to produce audit_marked_fast.xlsx.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    target_excel = _abs_path(target_excel)
    detect_rules_file = _abs_path(detect_rules_file)
    phase1_rules_file = _abs_path(phase1_rules_file) if phase1_rules_file else None
    out_excel_fast = _abs_path(out_excel_fast) if out_excel_fast else os.path.join(
        base_dir, "outputs", "audit_marked_fast.xlsx"
    )

    if not target_excel or not os.path.exists(target_excel):
        raise SystemExit(f"target_excel not found: {target_excel}")

    rules_file = phase1_rules_file or detect_rules_file
    if not rules_file or not os.path.exists(rules_file):
        raise SystemExit(f"rules file not found: {rules_file}")

    detect_rules = parse_detect_rules(rules_file)
    if not isinstance(detect_rules, dict) or len(detect_rules) == 0:
        raise SystemExit(f"detect_rules is empty or invalid: {rules_file}")

    audit_parser = build_arg_parser_rule_audit_mark_excel()
    resolved_window_height = (
        int(window_height)
        if window_height is not None
        else int(_safe_default(audit_parser, "window_height", 3))
    )
    resolved_window_width = (
        int(window_width)
        if window_width is not None
        else int(_safe_default(audit_parser, "window_width", 1))
    )
    resolved_tolerance = (
        float(tolerance)
        if tolerance is not None
        else float(_safe_default(audit_parser, "tolerance", 0.01))
    )

    audit_argv = [
        "--baseline-rules",
        rules_file,
        "--excel",
        target_excel,
        "--out-excel",
        out_excel_fast,
        "--window-height",
        str(resolved_window_height),
        "--window-width",
        str(resolved_window_width),
        "--tolerance",
        str(resolved_tolerance),
    ]
    if row_name:
        audit_argv.extend(["--row-name", str(row_name)])
    if strict_row_match:
        audit_argv.append("--strict-row-match")

    audit_args = audit_parser.parse_args(audit_argv)
    main_rule_audit_mark_excel(audit_args)

    print("=== ExcelStudio Fast-Mark Summary ===")
    print(f"target_excel: {target_excel}")
    print(f"rules_file: {rules_file}")
    print(f"out_excel_fast: {out_excel_fast}")
    return out_excel_fast

def parse_detect_rules(detect_rules_file):
    """
    Load detect_rules/discovered_rules json.
    """
    detect_rules = u.load_json(detect_rules_file)
    return detect_rules


def _get_run_math_rule_analysis_symbols():
    """
    Lazy import to avoid forcing heavy baseline-analysis dependencies
    when callers only need rules-only audit.
    """
    from dataProcess.run_math_rule_analysis import build_arg_parser as run_parser_builder
    from dataProcess.run_math_rule_analysis import main as run_main

    return run_parser_builder, run_main


def _safe_default(parser, key, fallback):
    try:
        value = parser.get_default(key)
    except Exception:
        return fallback
    if value is None:
        return fallback
    return value


def _abs_path(path_value):
    if path_value is None:
        return None
    p = os.path.expanduser(str(path_value))
    if not os.path.isabs(p):
        p = os.path.abspath(p)
    return p


def detect_target_with_rules(
    target_excel,
    detect_rules_file,
    out_excel=None,
    *,
    row_name=None,
    window_height=None,
    window_width=None,
    tolerance=None,
    strict_row_match=False,
):
    """
    Audit target Excel directly with an existing discovered_rules json.
    This path skips baseline rule discovery.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    audit_parser = build_arg_parser_rule_audit_mark_excel()

    target_excel = _abs_path(target_excel)
    detect_rules_file = _abs_path(detect_rules_file)
    out_excel = _abs_path(out_excel) if out_excel else os.path.join(base_dir, "outputs", "audit_marked.xlsx")

    if not os.path.exists(target_excel):
        raise SystemExit(f"target_excel not found: {target_excel}")
    if not os.path.exists(detect_rules_file):
        raise SystemExit(f"detect_rules_file not found: {detect_rules_file}")

    detect_rules = parse_detect_rules(detect_rules_file)
    if not isinstance(detect_rules, dict) or len(detect_rules) == 0:
        raise SystemExit(f"detect_rules is empty or invalid: {detect_rules_file}")

    resolved_window_height = (
        int(window_height)
        if window_height is not None
        else int(_safe_default(audit_parser, "window_height", 3))
    )
    resolved_window_width = (
        int(window_width)
        if window_width is not None
        else int(_safe_default(audit_parser, "window_width", 1))
    )
    resolved_tolerance = (
        float(tolerance)
        if tolerance is not None
        else float(_safe_default(audit_parser, "tolerance", 0.01))
    )

    audit_argv = [
        "--baseline-rules",
        detect_rules_file,
        "--excel",
        target_excel,
        "--out-excel",
        out_excel,
        "--window-height",
        str(resolved_window_height),
        "--window-width",
        str(resolved_window_width),
        "--tolerance",
        str(resolved_tolerance),
    ]
    if row_name:
        audit_argv.extend(["--row-name", str(row_name)])
    if strict_row_match:
        audit_argv.append("--strict-row-match")

    audit_args = audit_parser.parse_args(audit_argv)
    main_rule_audit_mark_excel(audit_args)

    print("=== ExcelStudio Rules-Only Summary ===")
    print(f"target_excel: {target_excel}")
    print(f"detect_rules_file: {detect_rules_file}")
    print(f"out_excel: {out_excel}")
    return out_excel


def detect_defect_simple_with_prompt_rules(
        detect_rules_file=None,
        target_excel=None,
        out_excel=None,
        **kwargs,
    ):
    """
    Convenience wrapper:
    - rules:   C:\\ML_HOME\\ExcelStudio\\prompts\\discovered_rules.json
    - target:  C:\\ML_HOME\\ExcelStudio\\defect_simple.xlsx
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    rules_path = detect_rules_file or os.path.join(base_dir, "prompts", "discovered_rules.json")
    target_path = target_excel or os.path.join(base_dir, "defect_simple_plank.xlsx")
    output_path = out_excel or os.path.join(base_dir, "outputs", "audit_marked_from_prompt_rules.xlsx")

    return detect_target_with_rules(
        target_excel=target_path,
        detect_rules_file=rules_path,
        out_excel=output_path,
        **kwargs,
    )


def main_rules_only(args):
    """
    CLI entry for auditing with an existing discovered_rules file only.
    Does not require baseline_excel.
    """
    detect_target_with_rules(
        target_excel=args.target_excel,
        detect_rules_file=args.detect_rules_file,
        out_excel=args.out_excel,
        row_name=args.row_name,
        window_height=args.window_height,
        window_width=args.window_width,
        tolerance=args.tolerance,
        strict_row_match=args.strict_row_match,
    )


def main(args, ret: Optional[Dict[str, Any]] = None):
    """
    1) Build/Load detect rules from baseline file (A).
    2) Use rules to audit target file (B) and output marked Excel.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    default_rules_path = os.path.join(base_dir, "outputs", "discovered_rules.json")

    baseline_excel = _abs_path(args.baseline_excel)
    target_excel = _abs_path(args.target_excel)
    out_excel = _abs_path(args.out_excel)
    detect_rules_file = _abs_path(args.detect_rules_file) if args.detect_rules_file else None

    if not os.path.exists(baseline_excel):
        raise SystemExit(f"baseline_excel not found: {baseline_excel}")
    if not os.path.exists(target_excel):
        raise SystemExit(f"target_excel not found: {target_excel}")

    # Step 1: If no detect rules file is given, generate it from baseline.
    if not detect_rules_file:
        run_parser_builder, run_math_rule_analysis_main = _get_run_math_rule_analysis_symbols()
        run_parser = run_parser_builder()
        run_argv = [
            baseline_excel,
            "--window-height", str(args.window_height),
            "--window-width", str(args.window_width),
            "--consistency-threshold", str(args.consistency_threshold),
            "--quick-scan-threshold", str(args.quick_scan_threshold),
        ]
        if args.row_name:
            run_argv.extend(["--start-loc-row-name", str(args.row_name)])
        if args.use_openai:
            run_argv.append("--use-openai")
        if args.openai_model:
            run_argv.extend(["--openai-model", str(args.openai_model)])

        run_args = run_parser.parse_args(run_argv)
        run_math_rule_analysis_main(run_args)
        detect_rules_file = default_rules_path
        # run_math_rule_analysis 內部完成 Phase1（與可選 Phase2）後，預設寫入 discover_rules／discovered_rules。
        # 若日後僅 Phase1 即寫出 checkpoint（例如 *_phase1.json），可在此讀取並傳入 apply_discovered_rules_fast_mark_target 的 phase1_rules_file。
    else:
        if not os.path.exists(detect_rules_file):
            raise SystemExit(f"detect_rules_file not found: {detect_rules_file}")

    # Step 2: Validate detect rules.
    detect_rules = parse_detect_rules(detect_rules_file)
    if not isinstance(detect_rules, dict) or len(detect_rules) == 0:
        raise SystemExit(f"detect_rules is empty or invalid: {detect_rules_file}")

    # Step 2b（快速標記骨架）：Phase1／規則 JSON 就緒後，另產 audit_marked_fast.xlsx；實作完成後在函式內寫入檔案。
    fast_mark_out = os.path.join(base_dir, "outputs", "audit_marked_fast.xlsx")
    phase1_rules_path = os.path.join(base_dir, "outputs", "discovered_rules_phase1.json")
    phase1_rules_file = phase1_rules_path if os.path.exists(phase1_rules_path) else None
    apply_discovered_rules_fast_mark_target(
        target_excel=target_excel,
        detect_rules_file=detect_rules_file,
        out_excel_fast=fast_mark_out,
        baseline_excel=baseline_excel,
        row_name=args.row_name,
        window_height=args.window_height,
        window_width=args.window_width,
        tolerance=args.tolerance,
        strict_row_match=args.strict_row_match,
        phase1_rules_file=phase1_rules_file,
    )

    # Step 3: Audit target excel with baseline rules and mark suspect cells.
    audit_parser = build_arg_parser_rule_audit_mark_excel()
    audit_argv = [
        "--baseline-rules", detect_rules_file,
        "--excel", target_excel,
        "--out-excel", out_excel,
        "--window-height", str(args.window_height),
        "--window-width", str(args.window_width),
        "--tolerance", str(args.tolerance),
    ]
    if args.row_name:
        audit_argv.extend(["--row-name", str(args.row_name)])
    if args.strict_row_match:
        audit_argv.append("--strict-row-match")

    audit_args = audit_parser.parse_args(audit_argv)
    main_rule_audit_mark_excel(audit_args, ret=ret)

    print("=== ExcelStudio Summary ===")
    print(f"baseline_excel: {baseline_excel}")
    print(f"target_excel: {target_excel}")
    print(f"detect_rules_file: {detect_rules_file}")
    print(f"out_excel: {out_excel}")


def build_arg_parser():
    audit_parser = build_arg_parser_rule_audit_mark_excel()
    run_parser_builder, _ = _get_run_math_rule_analysis_symbols()
    run_parser = run_parser_builder()

    parser = argparse.ArgumentParser(
        description="ExcelStudio: baseline rule discovery + target audit mark flow",
        usage=(
            "python ExcelStudio.py [rules-only ...]\n"
            "  full flow (default): python ExcelStudio.py BASELINE_EXCEL TARGET_EXCEL [options]\n"
            "  rules-only flow:     python ExcelStudio.py rules-only --detect-rules-file RULES --target-excel TARGET [options]"
        ),
    )

    parser.add_argument("baseline_excel", help="A file: baseline Excel")
    parser.add_argument("target_excel", help="B file: target Excel to audit")
    parser.add_argument(
        "-dr",
        "--detect-rules-file",
        default=None,
        help="Optional discovered_rules/detect_rules json. If omitted, generate from baseline_excel first.",
    )
    parser.add_argument(
        "-o",
        "--out-excel",
        default=os.path.join("outputs", "audit_marked.xlsx"),
        help="Output marked excel path",
    )

    parser.add_argument(
        "-rn",
        "--row-name",
        default=_safe_default(audit_parser, "row_name", None),
        help="Optional start_loc_row_name filter",
    )
    parser.add_argument(
        "-wh",
        "--window-height",
        type=int,
        default=_safe_default(run_parser, "window_height", _safe_default(audit_parser, "window_height", 3)),
    )
    parser.add_argument(
        "-ww",
        "--window-width",
        type=int,
        default=_safe_default(run_parser, "window_width", _safe_default(audit_parser, "window_width", 1)),
    )
    parser.add_argument(
        "-t",
        "--tolerance",
        type=float,
        default=_safe_default(audit_parser, "tolerance", 0.01),
    )
    parser.add_argument(
        "-srm",
        "--strict-row-match",
        action="store_true",
        default=False,
    )

    parser.add_argument(
        "--use-openai",
        action="store_true",
        default=_safe_default(run_parser, "use_openai", False),
    )
    parser.add_argument(
        "--openai-model",
        default=_safe_default(run_parser, "openai_model", "gpt35_chat"),
    )
    parser.add_argument(
        "--consistency-threshold",
        type=float,
        default=_safe_default(run_parser, "consistency_threshold", 0.8),
    )
    parser.add_argument(
        "--quick-scan-threshold",
        type=int,
        default=_safe_default(run_parser, "quick_scan_threshold", 3),
    )

    return parser


def build_arg_parser_rules_only():
    audit_parser = build_arg_parser_rule_audit_mark_excel()
    parser = argparse.ArgumentParser(
        description="ExcelStudio (rules-only): audit target Excel with an existing discovered_rules file",
    )
    # parser.add_argument("mode_prefix", help="B file: target Excel to audit")
    parser.add_argument("target_excel", help="B file: target Excel to audit")
    parser.add_argument(
        "detect_rules_file",
        help="Existing discovered_rules/detect_rules json",
    )
    parser.add_argument(
        "-o",
        "--out-excel",
        default=os.path.join("outputs", "audit_marked.xlsx"),
        help="Output marked excel path",
    )
    parser.add_argument(
        "-rn",
        "--row-name",
        default=_safe_default(audit_parser, "row_name", None),
        help="Optional start_loc_row_name filter",
    )
    parser.add_argument(
        "-wh",
        "--window-height",
        type=int,
        default=_safe_default(audit_parser, "window_height", 3),
    )
    parser.add_argument(
        "-ww",
        "--window-width",
        type=int,
        default=_safe_default(audit_parser, "window_width", 1),
    )
    parser.add_argument(
        "-t",
        "--tolerance",
        type=float,
        default=_safe_default(audit_parser, "tolerance", 0.01),
    )
    parser.add_argument(
        "-srm",
        "--strict-row-match",
        action="store_true",
        default=False,
        help="Only apply rules with matching row_name",
    )
    return parser


if __name__ == "__main__":
    # If the first arg is "rules-only", use the lightweight rules-only flow.
    if len(sys.argv) > 1 and sys.argv[1] == "UR":
        parser = build_arg_parser_rules_only()
        args = parser.parse_args(sys.argv[2:])
        main_rules_only(args)
    else:
        parser = build_arg_parser()
        args = parser.parse_args()
        main(args)
