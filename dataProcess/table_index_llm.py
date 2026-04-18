#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import json
import os
import argparse
from pathlib import Path

import requests
from dotenv import load_dotenv
import re
try:
    import pandas as pd
except Exception:  # pragma: no cover
    pd = None

# Ensure local ContextParser is importable
DATA_PROCESS_DIR = Path(__file__).resolve().parent
if str(DATA_PROCESS_DIR) not in sys.path:
    sys.path.insert(0, str(DATA_PROCESS_DIR))
from ContextParser import xlsx_parser as xp  # type: ignore

DEFAULT_FILE = r"C:\ML_HOME\ExcelStudio\correct_simple.xlsx"


def _is_numeric(value):
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return True
    text = str(value).strip()
    if text == "":
        return False
    text = text.replace(",", "")
    return bool(re.fullmatch(r"[-+]?\d+(\.\d+)?", text))


def _cell_type(value):
    if value is None:
        return "empty"
    text = str(value).strip()
    if text == "":
        return "empty"
    return "numeric" if _is_numeric(value) else "text"


def _compute_row_col_stats(values):
    rows = len(values)
    cols = max((len(r) for r in values), default=0)
    padded = [list(r) + [None] * (cols - len(r)) for r in values]

    col_stats = []
    for c in range(cols):
        non_empty = text_cnt = num_cnt = 0
        for r in range(rows):
            t = _cell_type(padded[r][c])
            if t == "empty":
                continue
            non_empty += 1
            if t == "text":
                text_cnt += 1
            else:
                num_cnt += 1
        col_stats.append({
            "index": c + 1,
            "non_empty": non_empty,
            "text": text_cnt,
            "numeric": num_cnt,
            "text_ratio": (text_cnt / non_empty) if non_empty else 0.0,
            "numeric_ratio": (num_cnt / non_empty) if non_empty else 0.0,
        })

    row_stats = []
    for r in range(rows):
        non_empty = text_cnt = num_cnt = 0
        for c in range(cols):
            t = _cell_type(padded[r][c])
            if t == "empty":
                continue
            non_empty += 1
            if t == "text":
                text_cnt += 1
            else:
                num_cnt += 1
        row_stats.append({
            "index": r + 1,
            "non_empty": non_empty,
            "text": text_cnt,
            "numeric": num_cnt,
            "text_ratio": (text_cnt / non_empty) if non_empty else 0.0,
            "numeric_ratio": (num_cnt / non_empty) if non_empty else 0.0,
        })
    return row_stats, col_stats


def _format_stats(row_stats, col_stats):
    lines = ["欄位統計（1-based index）:"]
    for s in col_stats:
        lines.append(
            f"C{s['index']}: non_empty={s['non_empty']} "
            f"text={s['text']} numeric={s['numeric']} "
            f"text_ratio={s['text_ratio']:.2f} numeric_ratio={s['numeric_ratio']:.2f}"
        )
    lines.append("列統計（1-based index）:")
    for s in row_stats:
        lines.append(
            f"R{s['index']}: non_empty={s['non_empty']} "
            f"text={s['text']} numeric={s['numeric']} "
            f"text_ratio={s['text_ratio']:.2f} numeric_ratio={s['numeric_ratio']:.2f}"
        )
    return "\n".join(lines)

def _get_cell(values, row_idx, col_idx):
    if row_idx < 1 or col_idx < 1:
        return None
    r = row_idx - 1
    c = col_idx - 1
    if r >= len(values):
        return None
    row = values[r]
    if c >= len(row):
        return None
    return row[c]


def crop_table_with_headers(
    values,
    row_header_index,
    col_header_index,
    data_row_start,
    data_row_end,
    data_col_start,
    data_col_end,
):
    """
    依「資料區」座標裁切表格，並保留對應欄/列標題。

    - row_header_index / col_header_index: 1-based，無則 -1
    - data_row_start/end, data_col_start/end: 1-based，針對「去掉標題後」的資料列/欄

    回傳：
    {
      "values": 裁切後含標題的二維陣列,
      "row_header_index": 新表格中的標題列位置 (1 或 -1),
      "col_header_index": 新表格中的標題欄位置 (1 或 -1)
    }
    """
    if values is None:
        return {"values": [], "row_header_index": -1, "col_header_index": -1}

    total_rows = len(values)
    total_cols = max((len(r) for r in values), default=0)
    rh = row_header_index if row_header_index and row_header_index > 0 else None
    ch = col_header_index if col_header_index and col_header_index > 0 else None

    row_indices = [i for i in range(1, total_rows + 1) if i != rh]
    col_indices = [i for i in range(1, total_cols + 1) if i != ch]

    if not row_indices or not col_indices:
        return {"values": [], "row_header_index": -1, "col_header_index": -1}

    rs = max(1, int(data_row_start))
    re = min(len(row_indices), int(data_row_end))
    cs = max(1, int(data_col_start))
    ce = min(len(col_indices), int(data_col_end))

    if rs > re or cs > ce:
        return {"values": [], "row_header_index": -1, "col_header_index": -1}

    data_rows_actual = row_indices[rs - 1 : re]
    data_cols_actual = col_indices[cs - 1 : ce]

    cropped = []
    if rh:
        header_row = []
        if ch:
            header_row.append(_get_cell(values, rh, ch))
        header_row.extend(_get_cell(values, rh, c) for c in data_cols_actual)
        cropped.append(header_row)

    for r in data_rows_actual:
        row_out = []
        if ch:
            row_out.append(_get_cell(values, r, ch))
        row_out.extend(_get_cell(values, r, c) for c in data_cols_actual)
        cropped.append(row_out)

    return {
        "values": cropped,
        "row_header_index": 1 if rh else -1,
        "col_header_index": 1 if ch else -1,
    }

def crop_table_with_headers_df(
    values,
    row_header_index,
    col_header_index,
    data_row_start,
    data_row_end,
    data_col_start,
    data_col_end,
):
    """
    使用 pandas 將欄列標題轉成 columns/index，再裁切資料區。

    - row_header_index / col_header_index: 1-based，無則 -1
    - data_row_start/end, data_col_start/end: 1-based，針對「去掉標題後」的資料列/欄
    回傳 dict: {"df": DataFrame, "row_header_index": int, "col_header_index": int}
    """
    if pd is None:
        raise RuntimeError("pandas 未安裝，無法使用 crop_table_with_headers_df")
    if values is None:
        return {"df": pd.DataFrame(), "row_header_index": -1, "col_header_index": -1}

    rows = len(values)
    cols = max((len(r) for r in values), default=0)
    if rows == 0 or cols == 0:
        return {"df": pd.DataFrame(), "row_header_index": -1, "col_header_index": -1}

    padded = [list(r) + [None] * (cols - len(r)) for r in values]
    df = pd.DataFrame(padded)

    rh = row_header_index if row_header_index and row_header_index > 0 else None
    ch = col_header_index if col_header_index and col_header_index > 0 else None

    row_indices = [i for i in range(1, rows + 1) if i != rh]
    col_indices = [i for i in range(1, cols + 1) if i != ch]
    if not row_indices or not col_indices:
        return {"df": pd.DataFrame(), "row_header_index": -1, "col_header_index": -1}

    rs = max(1, int(data_row_start))
    re = min(len(row_indices), int(data_row_end))
    cs = max(1, int(data_col_start))
    ce = min(len(col_indices), int(data_col_end))
    if rs > re or cs > ce:
        return {"df": pd.DataFrame(), "row_header_index": -1, "col_header_index": -1}

    data_rows_actual = row_indices[rs - 1 : re]
    data_cols_actual = col_indices[cs - 1 : ce]
    df_data = df.iloc[[r - 1 for r in data_rows_actual], [c - 1 for c in data_cols_actual]]

    if rh:
        header_row = df.iloc[rh - 1, [c - 1 for c in data_cols_actual]].tolist()
        df_data.columns = header_row
    if ch:
        header_col = df.iloc[[r - 1 for r in data_rows_actual], ch - 1].tolist()
        df_data.index = header_col

    return {
        "df": df_data,
        "row_header_index": 1 if rh else -1,
        "col_header_index": 1 if ch else -1,
    }

def _suggest_col_header(col_stats):
    if not col_stats:
        return None
    left = col_stats[0]
    others = col_stats[1:]
    if not others:
        return None
    others_non_empty = sum(c["non_empty"] for c in others)
    others_numeric = sum(c["numeric"] for c in others)
    others_numeric_ratio = (others_numeric / others_non_empty) if others_non_empty else 0.0
    if left["text_ratio"] >= 0.6 and others_numeric_ratio >= 0.6 and left["non_empty"] >= 3:
        return {
            "index": 1,
            "reason": "左欄多為文字描述，其他欄位多為數值，符合標題欄特徵",
        }
    return None


def _suggest_row_header(row_stats):
    if not row_stats:
        return None
    top = row_stats[:5]
    best = None
    top_non_empty = max((r["non_empty"] for r in top), default=1)
    for r in top:
        if r["non_empty"] < 3:
            continue
        score = r["text_ratio"] * 0.7 + (r["non_empty"] / max(1, top_non_empty)) * 0.3
        if best is None or score > best["score"]:
            best = {"index": r["index"], "score": score}
    if best:
        return {
            "index": best["index"],
            "reason": "前幾列中，文字比例與填值數較高，較像標題列",
        }
    return None


def _format_heuristic_hint(row_hint, col_hint):
    if not row_hint and not col_hint:
        return ""
    lines = ["啟發式建議（僅供參考，可忽略）："]
    if row_hint:
        lines.append(f"- row_header_candidate={row_hint['index']} ({row_hint['reason']})")
    if col_hint:
        lines.append(f"- col_header_candidate={col_hint['index']} ({col_hint['reason']})")
    return "\n".join(lines)


def build_prompt(table_text, stats_text, heuristic_text, heading_boost_situation=None):
    return (
        "你是表格結構判讀助手。以下表格內容為原始資料列，不含任何自動欄名列。\n"
        "請判斷：\n"
        "1) 哪一列是標題列（row header index，1-based），若沒有請回 -1。\n"
        "2) 哪一欄是標題欄（column header index，1-based），若沒有請回 -1。\n"
        "判斷提示：\n"
        "- 若大部分欄位是數值且只有最左側是文字描述，優先視最左欄為標題欄；若無明確訊號再回 -1。\n"
        "- 即使沒有強烈訊號，也要選出「最可能」的標題列（通常是最上方、含多個代號/單位的列）；只有完全無法判斷時才回 -1。\n"
        "請只輸出 JSON：{\n"
        "  \"row_header_index\": <int>,\n"
        "  \"col_header_index\": <int>,\n"
        "  \"row_reason\": \"為何選這一列當標題列，若無則說明沒有明顯標題列\",\n"
        "  \"col_reason\": \"為何選這一欄當標題欄，若無則說明沒有明顯標題欄\"\n"
        "}。\n"
        + (f"判斷提示：{heading_boost_situation}\n" if heading_boost_situation else "")
        + "\n統計摘要（可輔助判斷）：\n"
        + f"{stats_text}\n"
        + (f"\n{heuristic_text}\n" if heuristic_text else "")
        + "\n表格內容（TSV，每列一行）：\n"
        + f"{table_text}\n"
    )


def call_textprocessor(url, provider, model, prompt, max_tokens=256, temperature=0.2):
    payload = {
        "prompt": prompt,
        "provider": provider,
        "model": model,
        "max_tokens": max_tokens,
        "temperature": temperature,
    }
    resp = requests.post(url, json=payload, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    if "error" in data:
        raise RuntimeError(f"TextProcessor error: {data['error']}")
    content = data.get("output") or data.get("response") or data.get("message")
    return content, data


def main():
    load_dotenv()
    parser = argparse.ArgumentParser(description="LLM probe: detect row/column index for each table")
    parser.add_argument("-i", "--excel", default=DEFAULT_FILE, help="Excel file path")
    parser.add_argument("--url", default=os.getenv("TEXTPROCESSOR_URL", "http://10.1.3.127:6017/chat"))
    parser.add_argument("--provider", default=os.getenv("TEXTPROCESSOR_PROVIDER", "remote"))
    parser.add_argument("--model", default=os.getenv("TEXTPROCESSOR_MODEL", "remote8b"))
    parser.add_argument("--max-tokens", type=int, default=256)
    parser.add_argument("--temperature", type=float, default=0.2)
    parser.add_argument(
        "--heading-boost-situation",
        help="Heading boost situation",
        default="具有特殊名詞、每個位置都明顯不同、滿足上述條件的文字比數字更適合",
    )
    args = parser.parse_args()

    result = xp.process(
        file_path=args.excel,
        parse_mode="row",
        include_images=False,
        image_placeholder=False,
        enable_image_llm=False,
    )

    print("=== table segments ===")
    for seg in result["segments"]:
        if seg.get("segment_type") != "table":
            continue
        bbox = seg.get("table_bbox") or [
            seg.get("top_row"),
            seg.get("left_col"),
            seg.get("bottom_row"),
            seg.get("right_col"),
        ]
        shape = seg.get("table_df").shape if seg.get("table_df") is not None else None
        print(f"sheet={seg.get('sheet_name')} table_id={seg.get('table_id')} bbox={bbox} shape={shape}")

        values = seg.get("table_values") or []
        if not values:
            print("(empty table)")
            continue
        row_stats, col_stats = _compute_row_col_stats(values)
        stats_text = _format_stats(row_stats, col_stats)
        row_hint = _suggest_row_header(row_stats)
        col_hint = _suggest_col_header(col_stats)
        heuristic_text = _format_heuristic_hint(row_hint, col_hint)

        table_text = "\n".join(
            "\t".join("" if v is None else str(v) for v in row)
            for row in values
        )
        print(table_text)
        print("--- LLM output ---")
        try:
            output, raw = call_textprocessor(
                url=args.url,
                provider=args.provider,
                model=args.model,
                prompt=build_prompt(
                    table_text,
                    stats_text,
                    heuristic_text,
                    heading_boost_situation=args.heading_boost_situation,
                ),
                max_tokens=args.max_tokens,
                temperature=args.temperature,
            )
            print(output)
            try:
                parsed = json.loads(output)
            except Exception:
                m = re.search(r"\{.*\}", output, flags=re.S)
                parsed = json.loads(m.group(0)) if m else None
            if isinstance(parsed, dict) and col_hint and parsed.get("col_header_index") == -1:
                print(
                    f"[heuristic] suggest col_header_index={col_hint['index']} "
                    f"({col_hint['reason']})"
                )
            # 顯示完整回應以便查看 reason/解析細節
            try:
                print("[raw json]")
                print(json.dumps(raw, ensure_ascii=False, indent=2))
            except Exception:
                print(f"[raw str] {raw}")
        except Exception as e:
            print(f"[LLM error] {e}")
        print("-" * 40)

    print("\n=== stats ===")
    print(json.dumps(result["stats"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
