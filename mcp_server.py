# -*- coding: utf-8 -*-
"""
excel_mcp_server.py

將既有 FastAPI ExcelStudio API 包裝成 MCP server。
預設使用 stdio transport，可供 Cursor / Claude Desktop / MCP Gateway 呼叫。
也可使用 streamable-http transport。

安裝:
    pip install "mcp[cli]" httpx

啟動服務:
    python excel_mcp_server.py
    python excel_mcp_server.py --api-base http://10.1.3.127:7018 --transport stdio
    python excel_mcp_server.py --api-base http://10.1.3.127:7018 --transport streamable-http --host 10.1.3.127 --port 7019

使用:
1. 先啟動 api_server.py（FastAPI ExcelStudio API server）
       uvicorn api_server:app --host 10.1.3.127 --port 7018
2. 再啟動本 MCP server（對接 FastAPI backend）
3. Excel 相關工具會透過 /api/xlsx/command 呼叫 XlsxStdioServer
"""

from __future__ import annotations

import argparse
import os
from typing import Any, Dict, List, Optional

import httpx
from mcp.server.fastmcp import FastMCP


DEFAULT_API_BASE = os.getenv("EXCEL_API_BASE", "http://10.1.3.127:7018")
DEFAULT_TIMEOUT = float(os.getenv("EXCEL_API_TIMEOUT", "180"))

mcp = FastMCP(
    "excel",
    instructions=(
        "MCP server for ExcelStudio operations backed by an existing FastAPI Excel API. "
        "Use these tools to inspect, edit, audit, discover rules, and run xlsx stdio commands."
    ),
)

_API_BASE = DEFAULT_API_BASE.rstrip("/")
_TIMEOUT = DEFAULT_TIMEOUT


def set_runtime_config(api_base: str, timeout: float) -> None:
    global _API_BASE, _TIMEOUT
    _API_BASE = api_base.rstrip("/")
    _TIMEOUT = timeout


async def _request(
    method: str,
    path: str,
    *,
    params: Optional[Dict[str, Any]] = None,
    json_body: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    url = f"{_API_BASE}{path}"
    async with httpx.AsyncClient(timeout=_TIMEOUT) as client:
        resp = await client.request(method, url, params=params, json=json_body)

    try:
        data = resp.json()
    except Exception:
        resp.raise_for_status()
        return {
            "ok": False,
            "status_code": resp.status_code,
            "text": resp.text,
            "url": url,
        }

    if resp.is_success:
        return data

    return {
        "ok": False,
        "status_code": resp.status_code,
        "url": url,
        "error": data,
    }


def _clean_dict(d: Dict[str, Any]) -> Dict[str, Any]:
    return {k: v for k, v in d.items() if v is not None}


async def _xlsx_command(
    command: str,
    args: Optional[Dict[str, Any]] = None,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    body = {
        "command": command,
        "args": args or {},
        "request_id": request_id,
    }
    return await _request("POST", "/api/xlsx/command", json_body=body)


@mcp.tool()
async def health() -> Dict[str, Any]:
    """檢查 ExcelStudio API server 是否正常運行"""
    return await _request("GET", "/health")


@mcp.tool()
async def root_info() -> Dict[str, Any]:
    """檢查 ExcelStudio API server 的根資訊"""
    return await _request("GET", "/")


@mcp.tool()
async def clear_outputs() -> Dict[str, Any]:
    """清除 outputs 的輸出"""
    return await _request("POST", "/api/outputs/clear")


@mcp.tool()
async def get_rules_path() -> Dict[str, Any]:
    """獲取 rules 的路徑"""
    return await _request("GET", "/api/rules/path")


@mcp.tool()
async def get_discovered_rules() -> Dict[str, Any]:
    """獲取 discovered rules 的 JSON 資訊"""
    return await _request("GET", "/api/rules/discovered")


# ---------------------------------------------------------
# ExcelStudio 規則發現
# ---------------------------------------------------------

@mcp.tool()
async def rules_discover(
    baseline_excel: str,
    start_loc_row_name: Optional[str] = None,
    window_height: int = 3,
    window_width: int = 1,
    use_openai: bool = False,
    openai_model: str = "gpt35_chat",
    consistency_threshold: float = 0.8,
    quick_scan_threshold: int = 3,
    quick_scan_seed: Optional[int] = None,
    use_phase1_global_llm_batch: Optional[bool] = None,
    phase2_overlap_phase1_retro: Optional[bool] = None,
    step_rows: Optional[int] = None,
    step_cols: Optional[int] = None,
    degeneracy_min_nonzero_count: Optional[int] = None,
    degeneracy_min_distinct_nonzero: Optional[int] = None,
) -> Dict[str, Any]:
    """發現 baseline Excel 的規則"""
    body = _clean_dict({
        "baseline_excel": baseline_excel,
        "start_loc_row_name": start_loc_row_name,
        "window_height": window_height,
        "window_width": window_width,
        "use_openai": use_openai,
        "openai_model": openai_model,
        "consistency_threshold": consistency_threshold,
        "quick_scan_threshold": quick_scan_threshold,
        "quick_scan_seed": quick_scan_seed,
        "use_phase1_global_llm_batch": use_phase1_global_llm_batch,
        "phase2_overlap_phase1_retro": phase2_overlap_phase1_retro,
        "step_rows": step_rows,
        "step_cols": step_cols,
        "degeneracy_min_nonzero_count": degeneracy_min_nonzero_count,
        "degeneracy_min_distinct_nonzero": degeneracy_min_distinct_nonzero,
    })
    return await _request("POST", "/api/rules/discover", json_body=body)


@mcp.tool()
async def audit_excel(
    detect_rules_file: str,
    target_excel: str,
    out_excel: Optional[str] = None,
    row_name: Optional[str] = None,
    window_height: int = 3,
    window_width: int = 1,
    tolerance: float = 0.01,
    strict_row_match: bool = False,
) -> Dict[str, Any]:
    """稽核 target Excel 的規則"""
    body = _clean_dict({
        "detect_rules_file": detect_rules_file,
        "target_excel": target_excel,
        "out_excel": out_excel,
        "row_name": row_name,
        "window_height": window_height,
        "window_width": window_width,
        "tolerance": tolerance,
        "strict_row_match": strict_row_match,
    })
    return await _request("POST", "/api/audit", json_body=body)


@mcp.tool()
async def mark_fast(
    target_excel: str,
    detect_rules_file: str,
    out_excel_fast: Optional[str] = None,
    baseline_excel: Optional[str] = None,
    phase1_rules_file: Optional[str] = None,
    row_name: Optional[str] = None,
    window_height: int = 3,
    window_width: int = 1,
    tolerance: float = 0.01,
    strict_row_match: bool = False,
) -> Dict[str, Any]:
    """快速稽核 target Excel 的規則"""
    body = _clean_dict({
        "target_excel": target_excel,
        "detect_rules_file": detect_rules_file,
        "out_excel_fast": out_excel_fast,
        "baseline_excel": baseline_excel,
        "phase1_rules_file": phase1_rules_file,
        "row_name": row_name,
        "window_height": window_height,
        "window_width": window_width,
        "tolerance": tolerance,
        "strict_row_match": strict_row_match,
    })
    return await _request("POST", "/api/mark-fast", json_body=body)


@mcp.tool()
async def rules_only(
    target_excel: str,
    detect_rules_file: str,
    out_excel: Optional[str] = None,
    row_name: Optional[str] = None,
    window_height: int = 3,
    window_width: int = 1,
    tolerance: float = 0.01,
    strict_row_match: bool = False,
) -> Dict[str, Any]:
    """僅稽核 baseline 的規則"""
    body = _clean_dict({
        "target_excel": target_excel,
        "detect_rules_file": detect_rules_file,
        "out_excel": out_excel,
        "row_name": row_name,
        "window_height": window_height,
        "window_width": window_width,
        "tolerance": tolerance,
        "strict_row_match": strict_row_match,
    })
    return await _request("POST", "/api/rules-only", json_body=body)


@mcp.tool()
async def full_flow(
    baseline_excel: str,
    target_excel: str,
    detect_rules_file: Optional[str] = None,
    out_excel: Optional[str] = None,
    row_name: Optional[str] = None,
    window_height: int = 3,
    window_width: int = 1,
    tolerance: float = 0.01,
    strict_row_match: bool = False,
    use_openai: bool = False,
    openai_model: str = "gpt35_chat",
    consistency_threshold: float = 0.8,
    quick_scan_threshold: int = 3,
) -> Dict[str, Any]:
    """完整流程: 發現規則 + 稽核 target Excel"""
    body = _clean_dict({
        "baseline_excel": baseline_excel,
        "target_excel": target_excel,
        "detect_rules_file": detect_rules_file,
        "out_excel": out_excel,
        "row_name": row_name,
        "window_height": window_height,
        "window_width": window_width,
        "tolerance": tolerance,
        "strict_row_match": strict_row_match,
        "use_openai": use_openai,
        "openai_model": openai_model,
        "consistency_threshold": consistency_threshold,
        "quick_scan_threshold": quick_scan_threshold,
    })
    return await _request("POST", "/api/full-flow", json_body=body)


# ---------------------------------------------------------
# Task 管理
# ---------------------------------------------------------

@mcp.tool()
async def task_start(action: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    """啟動背景任務。action 可為 ruleDiscovery / audit / markFast / rulesOnly / fullFlow。"""
    return await _request(
        "POST",
        "/api/task/start",
        json_body={
            "action": action,
            "payload": payload,
        },
    )


@mcp.tool()
async def task_progress(task_id: Optional[str] = None) -> Dict[str, Any]:
    """查詢任務進度。"""
    params = {}
    if task_id:
        params["task_id"] = task_id
    return await _request("GET", "/api/task/progress", params=params)


@mcp.tool()
async def task_stop(task_id: Optional[str] = None) -> Dict[str, Any]:
    """停止指定任務（或目前執行中的任務）。"""
    return await _request(
        "POST",
        "/api/task/stop",
        json_body={"task_id": task_id},
    )


# ---------------------------------------------------------
# 通用 xlsx stdio command 工具
# ---------------------------------------------------------

@mcp.tool()
async def xlsx_command(
    command: str,
    args: Optional[Dict[str, Any]] = None,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """直接轉發 xlsx stdio command。"""
    return await _xlsx_command(command, args=args, request_id=request_id)


# ---------------------------------------------------------
# 常用 workbook / sheet / range MCP 工具
# ---------------------------------------------------------

@mcp.tool()
async def ping(request_id: Optional[str] = None) -> Dict[str, Any]:
    """xlsx stdio ping。"""
    return await _xlsx_command("ping", request_id=request_id)


@mcp.tool()
async def get_version(request_id: Optional[str] = None) -> Dict[str, Any]:
    """取得 xlsx stdio 版本資訊。"""
    return await _xlsx_command("get_version", request_id=request_id)


@mcp.tool()
async def list_commands(request_id: Optional[str] = None) -> Dict[str, Any]:
    """列出 xlsx stdio 支援的指令。"""
    return await _xlsx_command("list_commands", request_id=request_id)


@mcp.tool()
async def create_workbook(
    title: Optional[str] = None,
    sheets: Optional[List[str]] = None,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """建立新的 workbook session。"""
    return await _xlsx_command(
        "create_workbook",
        args=_clean_dict({
            "default_sheet_name": title,
            "initial_sheets": sheets,
        }),
        request_id=request_id,
    )


@mcp.tool()
async def open_workbook(
    path: str,
    read_only: bool = False,
    data_only: bool = False,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """開啟 workbook。"""
    return await _xlsx_command(
        "open_workbook",
        args={
            "path": path,
            "read_only": read_only,
            "data_only": data_only,
        },
        request_id=request_id,
    )


@mcp.tool()
async def save_workbook(
    workbook_id: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """儲存 workbook。"""
    return await _xlsx_command(
        "save_workbook",
        args={"workbook_id": workbook_id},
        request_id=request_id,
    )


@mcp.tool()
async def save_workbook_as(
    workbook_id: str,
    path: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """另存 workbook。"""
    return await _xlsx_command(
        "save_workbook_as",
        args={
            "workbook_id": workbook_id,
            "path": path,
        },
        request_id=request_id,
    )


@mcp.tool()
async def close_workbook(
    workbook_id: str,
    save_if_dirty: bool = False,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """關閉 workbook。"""
    return await _xlsx_command(
        "close_workbook",
        args={
            "workbook_id": workbook_id,
            "save_before_close": save_if_dirty,
        },
        request_id=request_id,
    )


@mcp.tool()
async def list_open_workbooks(request_id: Optional[str] = None) -> Dict[str, Any]:
    """列出目前開啟的 workbook。"""
    return await _xlsx_command("list_open_workbooks", request_id=request_id)


@mcp.tool()
async def get_workbook_info(
    workbook_id: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """取得 workbook 資訊。"""
    return await _xlsx_command(
        "get_workbook_info",
        args={"workbook_id": workbook_id},
        request_id=request_id,
    )


@mcp.tool()
async def list_sheets(
    workbook_id: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """列出 workbook 的工作表。"""
    return await _xlsx_command(
        "list_sheets",
        args={"workbook_id": workbook_id},
        request_id=request_id,
    )


@mcp.tool()
async def add_sheet(
    workbook_id: str,
    name: str,
    index: Optional[int] = None,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """新增工作表。"""
    return await _xlsx_command(
        "add_sheet",
        args=_clean_dict({
            "workbook_id": workbook_id,
            "title": name,
            "index": index,
        }),
        request_id=request_id,
    )


@mcp.tool()
async def rename_sheet(
    workbook_id: str,
    old_name: str,
    new_name: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """重新命名工作表。"""
    return await _xlsx_command(
        "rename_sheet",
        args={
            "workbook_id": workbook_id,
            "sheet": old_name,
            "new_title": new_name,
        },
        request_id=request_id,
    )


@mcp.tool()
async def delete_sheet(
    workbook_id: str,
    name: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """刪除工作表。"""
    return await _xlsx_command(
        "delete_sheet",
        args={
            "workbook_id": workbook_id,
            "sheet": name,
        },
        request_id=request_id,
    )


@mcp.tool()
async def read_range(
    workbook_id: str,
    sheet: str,
    range: str,
    mode: str = "matrix",
    include_formula: bool = False,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """讀取儲存格範圍。"""
    return await _xlsx_command(
        "read_range",
        args={
            "workbook_id": workbook_id,
            "sheet": sheet,
            "range": range,
            "mode": mode,
            "include_formula": include_formula,
        },
        request_id=request_id,
    )


@mcp.tool()
async def write_range(
    workbook_id: str,
    sheet: str,
    start_cell: str,
    values: List[List[Any]],
    overwrite: bool = True,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """寫入儲存格範圍。"""
    return await _xlsx_command(
        "write_range",
        args={
            "workbook_id": workbook_id,
            "sheet": sheet,
            "start_cell": start_cell,
            "values": values,
            "overwrite": overwrite,
        },
        request_id=request_id,
    )


@mcp.tool()
async def append_rows(
    workbook_id: str,
    sheet: str,
    rows: List[List[Any]],
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """附加列資料。"""
    return await _xlsx_command(
        "append_rows",
        args={
            "workbook_id": workbook_id,
            "sheet": sheet,
            "rows": rows,
        },
        request_id=request_id,
    )


@mcp.tool()
async def clear_range(
    workbook_id: str,
    sheet: str,
    range: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """清除儲存格範圍。"""
    return await _xlsx_command(
        "clear_range",
        args={
            "workbook_id": workbook_id,
            "sheet": sheet,
            "range": range,
        },
        request_id=request_id,
    )


@mcp.tool()
async def get_used_range(
    workbook_id: str,
    sheet: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """取得已使用範圍（used range）。"""
    return await _xlsx_command(
        "get_used_range",
        args={
            "workbook_id": workbook_id,
            "sheet": sheet,
        },
        request_id=request_id,
    )


@mcp.tool()
async def set_number_format(
    workbook_id: str,
    sheet: str,
    range: str,
    format_code: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """設定數字格式。"""
    return await _xlsx_command(
        "set_number_format",
        args={
            "workbook_id": workbook_id,
            "sheet": sheet,
            "range": range,
            "format_code": format_code,
        },
        request_id=request_id,
    )


@mcp.tool()
async def set_font_style(
    workbook_id: str,
    sheet: str,
    range: str,
    name: Optional[str] = None,
    size: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[str] = None,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """設定字型樣式。"""
    return await _xlsx_command(
        "set_font_style",
        args=_clean_dict({
            "workbook_id": workbook_id,
            "sheet": sheet,
            "range": range,
            "name": name,
            "size": size,
            "bold": bold,
            "italic": italic,
            "color": color,
        }),
        request_id=request_id,
    )


@mcp.tool()
async def set_fill_color(
    workbook_id: str,
    sheet: str,
    range: str,
    color: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """設定填滿顏色。"""
    return await _xlsx_command(
        "set_fill_color",
        args={
            "workbook_id": workbook_id,
            "sheet": sheet,
            "range": range,
            "color": color,
        },
        request_id=request_id,
    )


@mcp.tool()
async def set_alignment(
    workbook_id: str,
    sheet: str,
    range: str,
    horizontal: Optional[str] = None,
    vertical: Optional[str] = None,
    wrap_text: Optional[bool] = None,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """設定對齊方式。"""
    return await _xlsx_command(
        "set_alignment",
        args=_clean_dict({
            "workbook_id": workbook_id,
            "sheet": sheet,
            "range": range,
            "horizontal": horizontal,
            "vertical": vertical,
            "wrap_text": wrap_text,
        }),
        request_id=request_id,
    )


@mcp.tool()
async def set_formula(
    workbook_id: str,
    sheet: str,
    cell: str,
    formula: str,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """設定公式。"""
    return await _xlsx_command(
        "set_formula",
        args={
            "workbook_id": workbook_id,
            "sheet": sheet,
            "cell": cell,
            "formula": formula,
        },
        request_id=request_id,
    )


@mcp.tool()
async def create_table(
    workbook_id: str,
    sheet: str,
    range: str,
    table_name: str,
    style_name: Optional[str] = None,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """建立 Excel 表格（table）。"""
    return await _xlsx_command(
        "create_table",
        args=_clean_dict({
            "workbook_id": workbook_id,
            "sheet": sheet,
            "range": range,
            "name": table_name,
            "style_name": style_name,
        }),
        request_id=request_id,
    )


@mcp.tool()
async def export_csv(
    workbook_id: str,
    sheet: str,
    path: str,
    range: Optional[str] = None,
    encoding: Optional[str] = None,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """匯出 CSV。"""
    return await _xlsx_command(
        "export_csv",
        args=_clean_dict({
            "workbook_id": workbook_id,
            "sheet": sheet,
            "path": path,
            "range": range,
            "encoding": encoding,
        }),
        request_id=request_id,
    )


@mcp.tool()
async def export_json(
    workbook_id: str,
    sheet: str,
    path: str,
    range: Optional[str] = None,
    mode: Optional[str] = None,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """匯出 JSON。"""
    return await _xlsx_command(
        "export_json",
        args=_clean_dict({
            "workbook_id": workbook_id,
            "sheet": sheet,
            "path": path,
            "range": range,
            "mode": mode,
        }),
        request_id=request_id,
    )


@mcp.tool()
async def export_pdf(
    workbook_id: str,
    path: str,
    sheet: Optional[str] = None,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """export_pdf is not available in current xlsx stdio server."""
    return {
        "ok": False,
        "code": "COMMAND_NOT_SUPPORTED",
        "message": "xlsx stdio does not provide 'export_pdf'. Use export_csv/export_json or list_commands.",
        "workbook_id": workbook_id,
        "requested_path": path,
        "sheet": sheet,
        "request_id": request_id,
    }


@mcp.tool()
async def preview_group_aggregate(
    workbook_id: str,
    group_cols: List[str],
    value_col: str = "PGA",
    sheet: Optional[str] = None,
    range: Optional[str] = None,
    start_cell: Optional[str] = None,
    end_cell: Optional[str] = None,
    time_priority: Optional[List[str]] = None,
    include_empty_rows: bool = False,
    preview_limit: int = 10,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """預覽分組聚合結果，不寫回工作表。"""
    return await _xlsx_command(
        "preview_group_aggregate",
        args=_clean_dict({
            "workbook_id": workbook_id,
            "sheet": sheet,
            "range": range,
            "start_cell": start_cell,
            "end_cell": end_cell,
            "group_cols": group_cols,
            "value_col": value_col,
            "time_priority": time_priority,
            "include_empty_rows": include_empty_rows,
            "preview_limit": preview_limit,
        }),
        request_id=request_id,
    )


@mcp.tool()
async def group_aggregate(
    workbook_id: str,
    group_cols: List[str],
    value_col: str = "PGA",
    sheet: Optional[str] = None,
    range: Optional[str] = None,
    start_cell: Optional[str] = None,
    end_cell: Optional[str] = None,
    time_priority: Optional[List[str]] = None,
    include_empty_rows: bool = False,
    preview_limit: int = 10,
    write_sheet: Optional[str] = None,
    target_sheet: Optional[str] = None,
    replace_sheet: bool = False,
    clear_sheet: bool = True,
    start_cell_out: Optional[str] = None,
    return_records: bool = False,
    request_id: Optional[str] = None,
) -> Dict[str, Any]:
    """執行分組聚合，可選擇寫回工作表。"""
    args = _clean_dict({
        "workbook_id": workbook_id,
        "sheet": sheet,
        "range": range,
        "start_cell": start_cell,
        "end_cell": end_cell,
        "group_cols": group_cols,
        "value_col": value_col,
        "time_priority": time_priority,
        "include_empty_rows": include_empty_rows,
        "preview_limit": preview_limit,
        "write_sheet": write_sheet,
        "target_sheet": target_sheet,
        "replace_sheet": replace_sheet,
        "clear_sheet": clear_sheet,
        "return_records": return_records,
    })
    if start_cell_out is not None:
        args["start_cell"] = start_cell_out

    return await _xlsx_command(
        "group_aggregate",
        args=args,
        request_id=request_id,
    )

    
def main() -> None:
    parser = argparse.ArgumentParser(description="MCP server for ExcelStudio FastAPI backend")
    parser.add_argument(
        "--api-base",
        default=DEFAULT_API_BASE,
        help="ExcelStudio API base URL，例如 http://10.1.3.127:7018",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=DEFAULT_TIMEOUT,
        help="HTTP timeout seconds",
    )
    parser.add_argument(
        "--transport",
        default="stdio",
        choices=["stdio", "streamable-http", "sse"],
        help="MCP transport",
    )
    parser.add_argument(
        "--host",
        default="10.1.3.127",
        help="streamable-http / sse 綁定 host",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=7019,
        help="streamable-http / sse 綁定 port",
    )
    parser.add_argument(
        "--path",
        default="/mcp",
        help="streamable-http / sse 路徑",
    )
    args = parser.parse_args()

    set_runtime_config(args.api_base, args.timeout)

    if args.transport == "stdio":
        mcp.run(transport="stdio")
    elif args.transport == "streamable-http":
        mcp.run(transport="streamable-http", host=args.host, port=args.port, path=args.path)
    else:
        mcp.run(transport="sse", host=args.host, port=args.port, path=args.path)


if __name__ == "__main__":
    main()
