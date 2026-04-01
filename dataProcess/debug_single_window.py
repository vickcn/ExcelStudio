#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
單窗格偵測工具：指定 window_id 或座標，檢查單一窗格的 LLM 規律與驗證結果。
"""

import argparse
import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from universal_table_detector import UniversalTableDetector
from window_scanner import WindowScanner
from llm_math_rule_detector import LLMMathRuleDetector


def _resolve_excel_path(excel_file: str) -> Path:
    excel_path = Path(excel_file)
    if not excel_path.is_absolute():
        base_path = Path(__file__).parent.parent.parent
        excel_path = base_path / excel_file
    return excel_path


def _pick_window(
    windows: List[Dict[str, Any]],
    window_id: Optional[int],
    start_row: Optional[int],
    start_col: Optional[int],
) -> Optional[Dict[str, Any]]:
    if window_id is not None:
        for w in windows:
            if w.get('window_id') == window_id:
                return w
        return None
    if start_row is not None and start_col is not None:
        for w in windows:
            pos = w.get('position', {})
            if pos.get('start_row') == start_row and pos.get('start_col') == start_col:
                return w
        return None
    return None


def _print_window_info(window: Dict[str, Any]) -> None:
    info = {
        "window_id": window.get("window_id"),
        "excel_range": window.get("excel_range"),
        "shape": window.get("shape"),
        "position": window.get("position"),
        "start_loc": window.get("start_loc"),
        "start_loc_row_name": window.get("start_loc_row_name"),
        "start_loc_row_indicated": window.get("start_loc_row_indicated"),
    }
    print("\n=== Window Info ===")
    print(json.dumps(info, ensure_ascii=False, indent=2))


def main() -> None:
    parser = argparse.ArgumentParser(description="Debug single window for math rule detection")
    parser.add_argument("excel_file", nargs="?", default="correct_simple.xlsx",
                        help="Excel file path (default: correct_simple.xlsx)")
    parser.add_argument("--window-id", type=int, help="Target window_id from scanner")
    parser.add_argument("--start-row", type=int, help="Window start row (0-based)")
    parser.add_argument("--start-col", type=int, help="Window start col (0-based)")
    parser.add_argument("--window-height", type=int, default=3, help="Window height (default: 3)")
    parser.add_argument("--window-width", type=int, default=1, help="Window width (default: 1)")
    parser.add_argument("--start-loc-row-name", type=str, help="Optional start_loc_row_name filter")
    parser.add_argument("--use-openai", action="store_true",
                        help="Use TextProcessor OpenAI provider instead of remote8b")
    parser.add_argument("--openai-model", type=str, default="gpt35_chat",
                        help="OpenAI model alias (default: gpt35_chat)")
    args = parser.parse_args()

    if args.window_id is None and (args.start_row is None or args.start_col is None):
        print("請指定 --window-id 或 --start-row/--start-col")
        return

    excel_path = _resolve_excel_path(args.excel_file)
    if not excel_path.exists():
        print(f"找不到 Excel 檔案: {excel_path}")
        return

    detector = UniversalTableDetector()
    windows_scanner = WindowScanner((args.window_height, args.window_width))
    llm = LLMMathRuleDetector(prefer_local=not args.use_openai, openai_model=args.openai_model)

    print(f"excel_file: {excel_path}")
    print(f"window_shape: {(args.window_height, args.window_width)}")

    # 1) 偵測表格
    tables = detector.detect_tables_by_analysis(str(excel_path), table_mode="pure_numeric", use_llm=False)
    if not tables:
        print("未偵測到任何表格")
        return

    # 2) 擷取 DataFrame
    dataframes = detector.extract_dataframes(str(excel_path), tables)
    if not dataframes:
        print("無法擷取 DataFrame")
        return

    main_df = dataframes[0]
    print(f"main_df.shape: {main_df.shape}")

    # 3) 掃描視窗
    windows = windows_scanner.scan_dataframe(main_df)
    if args.start_loc_row_name is not None:
        windows = windows_scanner.filter_windows_by_start_loc_row(windows, args.start_loc_row_name)
    if not windows:
        print("未掃描到任何視窗")
        return

    # 4) 選擇視窗
    target = _pick_window(windows, args.window_id, args.start_row, args.start_col)
    if not target:
        print("找不到指定 window")
        return

    _print_window_info(target)
    prompt_data = windows_scanner.generate_prompt_data(target)
    print("\n=== Prompt Data ===")
    print(prompt_data.get("prompt_text"))

    if not prompt_data.get("has_numeric"):
        print("\n視窗無有效數值，略過 LLM 偵測")
        return

    # 5) LLM 偵測
    print("\n=== LLM Detection ===")
    llm_result = llm.detect_math_rules(
        prompt_data.get("values", []),
        prompt_data.get("row_names", []),
        prompt_data.get("column_names", []),
        args.use_openai
    )
    print(json.dumps(llm_result, ensure_ascii=False, indent=2))

    # 6) 規律驗證
    print("\n=== Validation ===")
    validation = llm.validate_rules(prompt_data.get("values", []), llm_result.get("rules", []))
    print(json.dumps(validation, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

