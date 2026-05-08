# api_server.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import csv
import io
import shutil
import os
import sys
import traceback
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

# ---------------------------------------------------------
# Path bootstrap
# ---------------------------------------------------------
CURRENT_DIR = Path(__file__).resolve().parent
DATA_PROCESS_DIR = CURRENT_DIR / "dataProcess"

if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))
if str(DATA_PROCESS_DIR) not in sys.path:
    sys.path.insert(0, str(DATA_PROCESS_DIR))

# ---------------------------------------------------------
# Existing project imports
# ---------------------------------------------------------
import ExcelStudio
from dataProcess import run_math_rule_analysis
from dataProcess import rule_audit_mark_excel
from xlsx_stdio import XlsxStdioServer


# ---------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------
app = FastAPI(
    title="ExcelStudio API Server",
    version="1.0.0",
    description="FastAPI driver for ExcelStudio full flow / rule discovery / audit / fast-mark",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    index_candidates = ["index.html", "index.htm", "excelstudio_dashboard.html"]
    base_url = str(request.base_url).rstrip("/")
    for name in index_candidates:
        index_path = templates_path / name
        if index_path.exists():
            return templates.TemplateResponse(
                name,
                {
                    "request": request,
                    "api_base_url": base_url,
                    "rules_file_path": str(_get_rules_file_path()),
                },
            )
    return HTMLResponse(content="<h1>ExcelStudio API Server</h1>", status_code=200)


# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def _abs_path(path_value: Optional[str]) -> Optional[str]:
    if path_value is None:
        return None
    p = Path(os.path.expanduser(str(path_value)))
    if not p.is_absolute():
        p = (CURRENT_DIR / p).resolve()
    return str(p)


def _exists(path_value: Optional[str]) -> bool:
    return bool(path_value) and Path(path_value).exists()


def _ensure_parent_dir(path_value: Optional[str]) -> Optional[str]:
    if not path_value:
        return path_value
    p = Path(path_value)
    p.parent.mkdir(parents=True, exist_ok=True)
    return str(p)


def _parser_supports(parser, option: str) -> bool:
    try:
        return option in parser._option_string_actions
    except Exception:
        return False


def _append_arg_if_supported(
    parser,
    argv: List[str],
    option: str,
    value: Any,
    *,
    is_flag: bool = False,
) -> None:
    if not _parser_supports(parser, option):
        return
    if is_flag:
        if bool(value):
            argv.append(option)
        return
    if value is None:
        return
    argv.extend([option, str(value)])


def _safe_outputs_dir() -> Path:
    out_dir = CURRENT_DIR / "outputs"
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def _safe_uploads_dir() -> Path:
    raw_dir = _get_config_value(
        config,
        "uploads_dir",
        "upload_dir",
        "uploads",
        default=str(CURRENT_DIR / "uploads"),
    )
    abs_dir = _abs_path(raw_dir) if raw_dir else str(CURRENT_DIR / "uploads")
    up_dir = Path(abs_dir)
    up_dir.mkdir(parents=True, exist_ok=True)
    return up_dir




def _load_config() -> Dict[str, Any]:
    config_path = CURRENT_DIR / "config.json"
    if not config_path.exists():
        return {}
    try:
        # 使用 utf-8-sig 以容忍 BOM，避免 JSONDecodeError
        with open(config_path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        return {}
    return {}


def _get_config_value(config: Dict[str, Any], *keys: str, default: Any = None) -> Any:
    for key in keys:
        if key in config:
            return config.get(key, default)
    return default


def _get_rules_file_path() -> Path:
    raw_path = _get_config_value(
        config,
        "rules_file_path",
        "detect_rules_file",
        "rules_file",
        "rules_path",
        default=str(_safe_outputs_dir() / "discovered_rules.json"),
    )
    abs_path = _abs_path(raw_path) if raw_path else None
    return Path(abs_path or str(_safe_outputs_dir() / "discovered_rules.json"))


config = _load_config()

templates_dir = _get_config_value(
    config,
    "templates_dir",
    "template_dir",
    "templates",
    default=str(CURRENT_DIR / "template"),
)
static_dir = _get_config_value(
    config,
    "static_dir",
    "static",
    default=str(CURRENT_DIR / "static"),
)

templates_path = Path(templates_dir).expanduser()
if not templates_path.is_absolute():
    templates_path = (CURRENT_DIR / templates_path).resolve()
static_path = Path(static_dir).expanduser()
if not static_path.is_absolute():
    static_path = (CURRENT_DIR / static_path).resolve()

templates = Jinja2Templates(directory=str(templates_path))
if static_path.exists():
    app.mount("/static", StaticFiles(directory=str(static_path)), name="static")


def _collect_rule_output_paths(suffix: str = "") -> Dict[str, str]:
    out_dir = _safe_outputs_dir()
    return {
        "discovered_rules": str(out_dir / f"discovered_rules{suffix}.json"),
        "rule_statistics": str(out_dir / f"rule_statistics{suffix}.json"),
        "observed": str(out_dir / f"observed{suffix}.json"),
        "analysis_result": str(out_dir / f"analysis_result{suffix}.json"),
        "timer": str(out_dir / f"timer{suffix}.json"),
    }


def _existing_only(d: Dict[str, str]) -> Dict[str, str]:
    return {k: v for k, v in d.items() if Path(v).exists()}


def _read_json_file(path: Path) -> Any:
    if not path.exists():
        raise _http_error(f"file not found: {path}")
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except Exception as exc:
        raise _http_error(f"failed to read json: {path}", status_code=500) from exc


def _clear_outputs_dir() -> Dict[str, Any]:
    out_dir = _safe_outputs_dir()
    removed: List[str] = []
    for child in out_dir.iterdir():
        try:
            if child.is_dir():
                shutil.rmtree(child, ignore_errors=True)
            else:
                child.unlink()
            removed.append(str(child))
        except Exception:
            # best-effort cleanup; continue
            pass
    return {"outputs_dir": str(out_dir), "removed": removed}


def _http_error(message: str, status_code: int = 400) -> HTTPException:
    return HTTPException(status_code=status_code, detail=message)


def _safe_filename(name: str) -> str:
    # minimal sanitization; keep Unicode, just strip path parts
    return Path(name).name


RULE_BLOB_STORE: Dict[str, Dict[str, Any]] = {}
RULE_BLOB_LOCK = threading.Lock()
RULE_BLOB_MAX_SIZE = 128

_xlsx_stdio_server: Optional[XlsxStdioServer] = None
_XLSX_STDIO_INIT_LOCK = threading.Lock()
_XLSX_STDIO_PROCESS_LOCK = threading.Lock()


def _get_xlsx_stdio_server() -> XlsxStdioServer:
    global _xlsx_stdio_server
    if _xlsx_stdio_server is not None:
        return _xlsx_stdio_server
    with _XLSX_STDIO_INIT_LOCK:
        if _xlsx_stdio_server is None:
            _xlsx_stdio_server = XlsxStdioServer()
        return _xlsx_stdio_server


def _store_rules_blob(rules: Any) -> str:
    rules_id = uuid.uuid4().hex
    with RULE_BLOB_LOCK:
        RULE_BLOB_STORE[rules_id] = {
            "created_at": _utc_now(),
            "rules": rules,
        }
        while len(RULE_BLOB_STORE) > RULE_BLOB_MAX_SIZE:
            oldest_key = next(iter(RULE_BLOB_STORE))
            RULE_BLOB_STORE.pop(oldest_key, None)
    return rules_id


def _get_rules_blob(rules_id: str) -> Optional[Any]:
    with RULE_BLOB_LOCK:
        item = RULE_BLOB_STORE.get(rules_id)
        return item.get("rules") if isinstance(item, dict) else None


def _coerce_cell_value(raw: Any) -> Any:
    if not isinstance(raw, str):
        return raw
    text = raw.strip()
    if text == "":
        return ""

    lowered = text.lower()
    if lowered in {"null", "none"}:
        return None
    if lowered == "true":
        return True
    if lowered == "false":
        return False

    try:
        return int(text)
    except Exception:
        pass
    try:
        return float(text)
    except Exception:
        pass
    return text


def _normalize_2d_table(value: Any, *, field_name: str) -> List[List[Any]]:
    source = value
    if isinstance(source, dict):
        if "table" in source:
            source = source.get("table")
        elif "data" in source:
            source = source.get("data")

    if not isinstance(source, list):
        raise _http_error(f"{field_name} must be a 2D array")

    table: List[List[Any]] = []
    for idx, row in enumerate(source):
        if not isinstance(row, (list, tuple)):
            raise _http_error(f"{field_name}[{idx}] must be an array")
        table.append([_coerce_cell_value(cell) for cell in row])
    if not table:
        raise _http_error(f"{field_name} cannot be empty")
    return table


def _parse_markdown_table(table_text: str) -> List[List[Any]]:
    lines = [line.strip() for line in table_text.splitlines() if line.strip()]
    parsed: List[List[Any]] = []
    for line in lines:
        if "|" not in line:
            continue
        stripped = line.strip().strip("|").strip()
        if not stripped:
            continue
        compact = stripped.replace(" ", "")
        if compact and all(ch in "-:|" for ch in compact):
            continue
        cells = [part.strip() for part in line.strip().strip("|").split("|")]
        parsed.append([_coerce_cell_value(cell) for cell in cells])
    return parsed


def _parse_delimited_table(table_text: str, delimiter: str) -> List[List[Any]]:
    rows: List[List[Any]] = []
    reader = csv.reader(io.StringIO(table_text), delimiter=delimiter)
    for row in reader:
        if not row:
            continue
        rows.append([_coerce_cell_value(cell) for cell in row])
    return rows


def _parse_table_text(table_text: str, table_format: str = "auto") -> List[List[Any]]:
    fmt = (table_format or "auto").strip().lower()
    if fmt not in {"auto", "csv", "tsv", "json", "markdown"}:
        raise _http_error("table_format must be one of auto/csv/tsv/json/markdown")
    if not isinstance(table_text, str) or not table_text.strip():
        raise _http_error("table text is required")

    stripped = table_text.strip()
    if fmt in {"auto", "json"}:
        if fmt == "json" or stripped[:1] in {"[", "{"}:
            try:
                return _normalize_2d_table(json.loads(stripped), field_name="table")
            except HTTPException:
                raise
            except Exception:
                if fmt == "json":
                    raise _http_error("invalid JSON table text")

    if fmt in {"auto", "markdown"}:
        md_rows = _parse_markdown_table(table_text)
        if md_rows:
            return md_rows
        if fmt == "markdown":
            raise _http_error("invalid markdown table text")

    if fmt == "tsv":
        tsv_rows = _parse_delimited_table(table_text, "\t")
        if not tsv_rows:
            raise _http_error("invalid TSV table text")
        return tsv_rows

    if fmt == "csv":
        csv_rows = _parse_delimited_table(table_text, ",")
        if not csv_rows:
            raise _http_error("invalid CSV table text")
        return csv_rows

    # auto fallback: tsv first, then csv
    if "\t" in table_text:
        auto_tsv = _parse_delimited_table(table_text, "\t")
        if auto_tsv:
            return auto_tsv
    auto_csv = _parse_delimited_table(table_text, ",")
    if auto_csv:
        return auto_csv
    raise _http_error("unable to parse table text")


def _resolve_rules_payload(
    *,
    rules_id: Optional[str],
    rules_json: Optional[str],
    rules: Any,
    detect_rules_file: Optional[str],
) -> Dict[str, Any]:
    if rules is not None:
        payload = rules
        source = "rules"
    elif rules_json:
        try:
            payload = json.loads(rules_json)
        except Exception:
            raise _http_error("invalid rules_json")
        source = "rules_json"
    elif rules_id:
        payload = _get_rules_blob(rules_id)
        if payload is None:
            raise _http_error(f"rules_id not found: {rules_id}", status_code=404)
        source = "rules_id"
    elif detect_rules_file:
        abs_path = _abs_path(detect_rules_file)
        if not _exists(abs_path):
            raise _http_error(f"detect_rules_file not found: {abs_path}")
        payload = _read_json_file(Path(abs_path))
        source = "detect_rules_file"
    else:
        raise _http_error(
            "one of rules_id/rules_json/rules/detect_rules_file is required",
        )

    if not isinstance(payload, (dict, list)):
        raise _http_error("rules payload must be a JSON object or array")
    return {"rules": payload, "source": source}


def _run_audit_json_with_rules(
    *,
    target_table: List[List[Any]],
    target_sheet_name: Optional[str],
    row_name: Optional[str],
    window_height: int,
    window_width: int,
    tolerance: float,
    strict_row_match: bool,
    rules_payload: Any,
    rules_source: str,
) -> Dict[str, Any]:
    tmp_dir = _safe_outputs_dir() / "agent_rules_cache"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_path = tmp_dir / f"rules_{uuid.uuid4().hex}.json"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(rules_payload, f, ensure_ascii=False, indent=2)

    try:
        audit_req = AuditJsonRequest(
            detect_rules_file=str(tmp_path),
            target_table=target_table,
            target_sheet_name=target_sheet_name,
            row_name=row_name,
            window_height=window_height,
            window_width=window_width,
            tolerance=tolerance,
            strict_row_match=strict_row_match,
        )
        result = _run_audit_json(audit_req)
    finally:
        try:
            tmp_path.unlink()
        except Exception:
            pass

    result["rules_source"] = rules_source
    result["detect_rules_file"] = None
    return result


def _build_suspect_summary_text(result: Dict[str, Any]) -> str:
    suspect_sheet = result.get("suspect_sheet")
    suspect_cells = result.get("suspect_cells") or []
    addresses = [str(item.get("address")) for item in suspect_cells if isinstance(item, dict) and item.get("address")]
    preview = ", ".join(addresses[:20]) if addresses else "none"
    return (
        f"suspect_sheet={suspect_sheet or 'unknown'}; "
        f"suspect_cells_count={len(suspect_cells)}; "
        f"suspect_cells_preview={preview}"
    )


def _estimate_rules_count(payload: Any) -> int:
    if isinstance(payload, list):
        return len(payload)
    if isinstance(payload, dict):
        for key in ("rules", "passed_rules", "discovered_rules", "baseline_rules"):
            value = payload.get(key)
            if isinstance(value, list):
                return len(value)
    return 0


# ---------------------------------------------------------
# Task progress tracking
# ---------------------------------------------------------
class _TaskCancelled(Exception):
    pass


TASK_LOCK = threading.Lock()
TASK_STATE: Dict[str, Any] = {
    "task_id": None,
    "status": "idle",
    "action": None,
    "progress": 0,
    "message": "",
    "result": None,
    "error": None,
    "started_at": None,
    "updated_at": None,
    "cancel_event": None,
}


def _utc_now() -> str:
    return datetime.utcnow().isoformat()


def _task_public_state() -> Dict[str, Any]:
    with TASK_LOCK:
        return {
            "task_id": TASK_STATE.get("task_id"),
            "status": TASK_STATE.get("status"),
            "action": TASK_STATE.get("action"),
            "progress": TASK_STATE.get("progress", 0),
            "message": TASK_STATE.get("message"),
            "result": TASK_STATE.get("result"),
            "error": TASK_STATE.get("error"),
            "started_at": TASK_STATE.get("started_at"),
            "updated_at": TASK_STATE.get("updated_at"),
        }


def _set_task_state(**kwargs: Any) -> None:
    with TASK_LOCK:
        TASK_STATE.update(kwargs)
        TASK_STATE["updated_at"] = _utc_now()


def _set_task_progress(progress: int, message: str) -> None:
    _set_task_state(progress=int(progress), message=message)


# ---------------------------------------------------------
# Request models
# ---------------------------------------------------------
class CommonWindowOptions(BaseModel):
    row_name: Optional[str] = Field(default=None, description="start_loc_row_name filter")
    window_height: int = 3
    window_width: int = 1
    tolerance: float = 0.01
    strict_row_match: bool = False


class RuleDiscoveryRequest(BaseModel):
    baseline_excel: str
    start_loc_row_name: Optional[str] = None
    window_height: int = 3
    window_width: int = 1

    use_openai: bool = False
    openai_model: str = "gpt35_chat"
    consistency_threshold: float = 0.8
    quick_scan_threshold: int = 3

    # optional advanced params; only passed if parser supports them
    quick_scan_seed: Optional[int] = None
    use_phase1_global_llm_batch: Optional[bool] = None
    phase2_overlap_phase1_retro: Optional[bool] = None
    step_rows: Optional[int] = None
    step_cols: Optional[int] = None
    degeneracy_min_nonzero_count: Optional[int] = None
    degeneracy_min_distinct_nonzero: Optional[int] = None


class RuleDiscoveryJsonRequest(BaseModel):
    baseline_table: List[List[Any]]
    baseline_sheet_name: Optional[str] = None
    start_loc_row_name: Optional[str] = None
    window_height: int = 3
    window_width: int = 1

    use_openai: bool = False
    openai_model: str = "gpt35_chat"
    consistency_threshold: float = 0.8
    quick_scan_threshold: int = 3

    quick_scan_seed: Optional[int] = None
    use_phase1_global_llm_batch: Optional[bool] = None
    phase2_overlap_phase1_retro: Optional[bool] = None
    step_rows: Optional[int] = None
    step_cols: Optional[int] = None
    degeneracy_min_nonzero_count: Optional[int] = None
    degeneracy_min_distinct_nonzero: Optional[int] = None


class AuditRequest(CommonWindowOptions):
    detect_rules_file: str
    target_excel: str
    out_excel: Optional[str] = None


class AuditJsonRequest(CommonWindowOptions):
    detect_rules_file: str
    target_table: List[List[Any]]
    target_sheet_name: Optional[str] = None


class FastMarkRequest(CommonWindowOptions):
    target_excel: str
    detect_rules_file: str
    out_excel_fast: Optional[str] = None
    baseline_excel: Optional[str] = None
    phase1_rules_file: Optional[str] = None


class FullFlowRequest(CommonWindowOptions):
    baseline_excel: str
    target_excel: str
    detect_rules_file: Optional[str] = None
    out_excel: Optional[str] = None

    use_openai: bool = False
    openai_model: str = "gpt35_chat"
    consistency_threshold: float = 0.8
    quick_scan_threshold: int = 3


class FullFlowJsonRequest(CommonWindowOptions):
    baseline_table: List[List[Any]]
    baseline_sheet_name: Optional[str] = None
    target_table: List[List[Any]]
    target_sheet_name: Optional[str] = None
    detect_rules_file: Optional[str] = None

    use_openai: bool = False
    openai_model: str = "gpt35_chat"
    consistency_threshold: float = 0.8
    quick_scan_threshold: int = 3

    quick_scan_seed: Optional[int] = None
    use_phase1_global_llm_batch: Optional[bool] = None
    phase2_overlap_phase1_retro: Optional[bool] = None
    step_rows: Optional[int] = None
    step_cols: Optional[int] = None
    degeneracy_min_nonzero_count: Optional[int] = None
    degeneracy_min_distinct_nonzero: Optional[int] = None


class RuleDiscoveryAgentRequest(BaseModel):
    baseline_text: str = Field(..., description="Table text in csv/tsv/json/markdown")
    table_format: str = Field(default="auto", description="auto/csv/tsv/json/markdown")
    baseline_sheet_name: Optional[str] = None
    start_loc_row_name: Optional[str] = None
    window_height: int = 3
    window_width: int = 1

    use_openai: bool = False
    openai_model: str = "gpt35_chat"
    consistency_threshold: float = 0.8
    quick_scan_threshold: int = 3

    quick_scan_seed: Optional[int] = None
    use_phase1_global_llm_batch: Optional[bool] = None
    phase2_overlap_phase1_retro: Optional[bool] = None
    step_rows: Optional[int] = None
    step_cols: Optional[int] = None
    degeneracy_min_nonzero_count: Optional[int] = None
    degeneracy_min_distinct_nonzero: Optional[int] = None

    store_rules: bool = True
    return_rules_json: bool = True


class AuditAgentRequest(CommonWindowOptions):
    target_text: str = Field(..., description="Table text in csv/tsv/json/markdown")
    table_format: str = Field(default="auto", description="auto/csv/tsv/json/markdown")
    target_sheet_name: Optional[str] = None

    rules_id: Optional[str] = None
    rules_json: Optional[str] = None
    rules: Optional[Any] = None
    detect_rules_file: Optional[str] = None

    store_rules: bool = False
    return_rules_json: bool = False


class FullFlowAgentRequest(CommonWindowOptions):
    baseline_text: str = Field(..., description="Table text in csv/tsv/json/markdown")
    target_text: str = Field(..., description="Table text in csv/tsv/json/markdown")
    baseline_table_format: str = Field(default="auto", description="auto/csv/tsv/json/markdown")
    target_table_format: str = Field(default="auto", description="auto/csv/tsv/json/markdown")
    baseline_sheet_name: Optional[str] = None
    target_sheet_name: Optional[str] = None
    start_loc_row_name: Optional[str] = None

    use_openai: bool = False
    openai_model: str = "gpt35_chat"
    consistency_threshold: float = 0.8
    quick_scan_threshold: int = 3

    quick_scan_seed: Optional[int] = None
    use_phase1_global_llm_batch: Optional[bool] = None
    phase2_overlap_phase1_retro: Optional[bool] = None
    step_rows: Optional[int] = None
    step_cols: Optional[int] = None
    degeneracy_min_nonzero_count: Optional[int] = None
    degeneracy_min_distinct_nonzero: Optional[int] = None

    rules_id: Optional[str] = None
    rules_json: Optional[str] = None
    rules: Optional[Any] = None
    detect_rules_file: Optional[str] = None
    force_discovery: bool = False

    store_rules: bool = True
    return_rules_json: bool = True


class RulesOnlyRequest(CommonWindowOptions):
    target_excel: str
    detect_rules_file: str
    out_excel: Optional[str] = None


class TaskStartRequest(BaseModel):
    action: str
    payload: Dict[str, Any] = Field(default_factory=dict)


class TaskStopRequest(BaseModel):
    task_id: Optional[str] = None


class XlsxStdioCommandRequest(BaseModel):
    command: str
    args: Dict[str, Any] = Field(default_factory=dict)
    request_id: Optional[str] = None


class PreviewGroupAggregateRequest(BaseModel):
    workbook_id: str
    sheet: Optional[str] = None
    range: Optional[str] = None
    start_cell: Optional[str] = None
    end_cell: Optional[str] = None
    group_cols: List[str]
    value_col: str = "PGA"
    time_priority: Optional[List[str]] = None
    include_empty_rows: bool = False
    preview_limit: int = 10
    request_id: Optional[str] = None


class GroupAggregateRequest(BaseModel):
    workbook_id: str
    sheet: Optional[str] = None
    range: Optional[str] = None
    start_cell: Optional[str] = None
    end_cell: Optional[str] = None
    group_cols: List[str]
    value_col: str = "PGA"
    time_priority: Optional[List[str]] = None
    include_empty_rows: bool = False
    preview_limit: int = 10
    write_sheet: Optional[str] = None
    target_sheet: Optional[str] = None
    replace_sheet: bool = False
    clear_sheet: bool = True
    start_cell_out: Optional[str] = None
    return_records: bool = False
    request_id: Optional[str] = None


class CaptureRangeImageRequest(BaseModel):
    workbook_id: str
    sheet: Optional[str] = None
    center_cell: str
    output_path: str
    up: int = 8
    down: int = 8
    left: int = 8
    right: int = 8
    request_id: Optional[str] = None


class CaptureRangeFigureTask(BaseModel):
    center_cell: str
    up: int = 8
    down: int = 8
    left: int = 8
    right: int = 8
    label: Optional[str] = None


class CaptureRangeFigureRequest(BaseModel):
    workbook_id: str
    sheet: Optional[str] = None
    output_path: str
    tasks: List[CaptureRangeFigureTask]
    ncols: int = 3
    panel_gap: int = 16
    panel_padding: int = 8
    request_id: Optional[str] = None

def _run_xlsx_stdio_command(req: XlsxStdioCommandRequest) -> Dict[str, Any]:
    payload: Dict[str, Any] = {
        "command": req.command,
        "args": dict(req.args) if req.args else {},
        "request_id": req.request_id,
    }
    server = _get_xlsx_stdio_server()
    with _XLSX_STDIO_PROCESS_LOCK:
        return server.process_payload(payload)


# ---------------------------------------------------------
# Core runners
# ---------------------------------------------------------
def _run_rule_discovery(req: RuleDiscoveryRequest) -> Dict[str, Any]:
    baseline_excel = _abs_path(req.baseline_excel)
    if not _exists(baseline_excel):
        raise _http_error(f"baseline_excel not found: {baseline_excel}")

    parser = run_math_rule_analysis.build_arg_parser()
    argv: List[str] = [baseline_excel]

    _append_arg_if_supported(parser, argv, "--start-loc-row-name", req.start_loc_row_name)
    _append_arg_if_supported(parser, argv, "--window-height", req.window_height)
    _append_arg_if_supported(parser, argv, "--window-width", req.window_width)
    _append_arg_if_supported(parser, argv, "--consistency-threshold", req.consistency_threshold)
    _append_arg_if_supported(parser, argv, "--quick-scan-threshold", req.quick_scan_threshold)
    _append_arg_if_supported(parser, argv, "--use-openai", req.use_openai, is_flag=True)
    _append_arg_if_supported(parser, argv, "--openai-model", req.openai_model)

    _append_arg_if_supported(parser, argv, "--quick-scan-seed", req.quick_scan_seed)
    _append_arg_if_supported(
        parser,
        argv,
        "--use-phase1-global-llm-batch",
        req.use_phase1_global_llm_batch,
        is_flag=True,
    )
    _append_arg_if_supported(
        parser,
        argv,
        "--phase2-overlap-phase1-retro",
        req.phase2_overlap_phase1_retro,
        is_flag=True,
    )
    _append_arg_if_supported(parser, argv, "--step-rows", req.step_rows)
    _append_arg_if_supported(parser, argv, "--step-cols", req.step_cols)
    _append_arg_if_supported(
        parser,
        argv,
        "--degeneracy-min-nonzero-count",
        req.degeneracy_min_nonzero_count,
    )
    _append_arg_if_supported(
        parser,
        argv,
        "--degeneracy-min-distinct-nonzero",
        req.degeneracy_min_distinct_nonzero,
    )

    args = parser.parse_args(argv)
    run_math_rule_analysis.main(args)

    outputs = {}
    outputs.update(_existing_only(_collect_rule_output_paths()))
    outputs.update(_existing_only(_collect_rule_output_paths("_phase1")))
    outputs.update(_existing_only(_collect_rule_output_paths("_phase2_final")))

    return {
        "success": True,
        "mode": "rule_discovery",
        "baseline_excel": baseline_excel,
        "request_argv": argv,
        "outputs": outputs,
    }


def _build_rule_analyzer_from_req(req: Any) -> run_math_rule_analysis.MathRuleAnalyzer:
    window_shape = (req.window_height, req.window_width)
    step_rows = req.step_rows if req.step_rows is not None else 1
    step_cols = req.step_cols if req.step_cols is not None else 1
    step_size = (step_rows, step_cols)
    min_nonzero = (
        req.degeneracy_min_nonzero_count
        if req.degeneracy_min_nonzero_count is not None
        else run_math_rule_analysis.DEFAULT_MIN_NONZERO_COUNT
    )
    min_distinct = (
        req.degeneracy_min_distinct_nonzero
        if req.degeneracy_min_distinct_nonzero is not None
        else run_math_rule_analysis.DEFAULT_MIN_DISTINCT_NONZERO
    )
    analyzer = run_math_rule_analysis.MathRuleAnalyzer(
        use_openai=req.use_openai,
        window_shape=window_shape,
        step_size=step_size,
        openai_model=req.openai_model,
        consistency_threshold=req.consistency_threshold,
        quick_scan_threshold=req.quick_scan_threshold,
        quick_scan_seed=req.quick_scan_seed,
        use_phase1_global_llm_batch=(
            req.use_phase1_global_llm_batch
            if req.use_phase1_global_llm_batch is not None
            else True
        ),
        phase2_overlap_phase1_retro=(
            req.phase2_overlap_phase1_retro
            if req.phase2_overlap_phase1_retro is not None
            else True
        ),
        degeneracy_min_nonzero_count=min_nonzero,
        degeneracy_min_distinct_nonzero=min_distinct,
    )
    return analyzer


def _run_rule_discovery_json(req: RuleDiscoveryJsonRequest) -> Dict[str, Any]:
    analyzer = _build_rule_analyzer_from_req(req)
    result = analyzer.analyze_sheetTable(
        req.baseline_table,
        start_loc_row_name=req.start_loc_row_name,
        sheet_name=req.baseline_sheet_name,
    )

    outputs = {}
    detect_rules_file = None
    if result.get("success"):
        detect_rules_file = analyzer.save_rule_summary(
            result,
            consistency_threshold=req.consistency_threshold,
        )
        analyzer.save_to_observed_json(result)
        analyzer.save_analysis_snapshot(result)
        analyzer.save_timer()
        if req.quick_scan_threshold > 0:
            analyzer.save_rule_summary(
                result,
                consistency_threshold=req.consistency_threshold,
                file_suffix="_phase2_final",
                print_summary=False,
            )
            analyzer.save_to_observed_json(result, file_suffix="_phase2_final")
            analyzer.save_analysis_snapshot(result, file_suffix="_phase2_final")
            analyzer.save_timer(file_suffix="_phase2_final")

        outputs.update(_existing_only(_collect_rule_output_paths()))
        outputs.update(_existing_only(_collect_rule_output_paths("_phase1")))
        outputs.update(_existing_only(_collect_rule_output_paths("_phase2_final")))

    return {
        "success": bool(result.get("success")),
        "mode": "rule_discovery_json",
        "baseline_sheet_name": req.baseline_sheet_name,
        "detect_rules_file": detect_rules_file,
        "outputs": outputs,
        "analysis": result,
    }


def _run_audit(req: AuditRequest) -> Dict[str, Any]:
    detect_rules_file = _abs_path(req.detect_rules_file)
    target_excel = _abs_path(req.target_excel)
    out_excel = _abs_path(req.out_excel) if req.out_excel else str(_safe_outputs_dir() / "audit_marked.xlsx")
    out_excel = _ensure_parent_dir(out_excel)

    if not _exists(detect_rules_file):
        raise _http_error(f"detect_rules_file not found: {detect_rules_file}")
    if not _exists(target_excel):
        raise _http_error(f"target_excel not found: {target_excel}")

    parser = rule_audit_mark_excel.build_arg_parser()
    argv: List[str] = [
        "--baseline-rules", detect_rules_file,
        "--excel", target_excel,
        "--out-excel", out_excel,
        "--window-height", str(req.window_height),
        "--window-width", str(req.window_width),
        "--tolerance", str(req.tolerance),
    ]
    if req.row_name:
        argv.extend(["--row-name", req.row_name])
    if req.strict_row_match:
        argv.append("--strict-row-match")

    args = parser.parse_args(argv)
    mark_excel_ret = {}
    rule_audit_mark_excel.main(args, ret=mark_excel_ret)

    return {
        "success": True,
        "mode": "audit",
        "detect_rules_file": detect_rules_file,
        "target_excel": target_excel,
        "out_excel": out_excel,
        "request_argv": argv,
        "mark_summary": mark_excel_ret,
        "suspect_sheet": mark_excel_ret.get("suspect_sheet") if isinstance(mark_excel_ret, dict) else None,
        "suspect_cells": mark_excel_ret.get("suspect_cells", []) if isinstance(mark_excel_ret, dict) else [],
    }


def _run_audit_json(req: AuditJsonRequest) -> Dict[str, Any]:
    detect_rules_file = _abs_path(req.detect_rules_file)
    if not _exists(detect_rules_file):
        raise _http_error(f"detect_rules_file not found: {detect_rules_file}")

    mark_excel_ret = rule_audit_mark_excel.audit_table_data(
        req.target_table,
        Path(detect_rules_file),
        sheet_name=req.target_sheet_name,
        row_name=req.row_name,
        window_height=req.window_height,
        window_width=req.window_width,
        tolerance=req.tolerance,
        strict_row_match=req.strict_row_match,
    )

    return {
        "success": True,
        "mode": "audit_json",
        "detect_rules_file": detect_rules_file,
        "target_sheet_name": req.target_sheet_name,
        "mark_summary": mark_excel_ret,
        "suspect_sheet": mark_excel_ret.get("suspect_sheet") if isinstance(mark_excel_ret, dict) else None,
        "suspect_cells": mark_excel_ret.get("suspect_cells", []) if isinstance(mark_excel_ret, dict) else [],
    }


def _run_fast_mark(req: FastMarkRequest) -> Dict[str, Any]:
    target_excel = _abs_path(req.target_excel)
    detect_rules_file = _abs_path(req.detect_rules_file)
    baseline_excel = _abs_path(req.baseline_excel)
    phase1_rules_file = _abs_path(req.phase1_rules_file)
    out_excel_fast = _abs_path(req.out_excel_fast) if req.out_excel_fast else str(_safe_outputs_dir() / "audit_marked_fast.xlsx")
    out_excel_fast = _ensure_parent_dir(out_excel_fast)

    if not _exists(target_excel):
        raise _http_error(f"target_excel not found: {target_excel}")
    if not _exists(detect_rules_file):
        raise _http_error(f"detect_rules_file not found: {detect_rules_file}")
    if baseline_excel and not _exists(baseline_excel):
        raise _http_error(f"baseline_excel not found: {baseline_excel}")
    if phase1_rules_file and not _exists(phase1_rules_file):
        raise _http_error(f"phase1_rules_file not found: {phase1_rules_file}")

    out_path = ExcelStudio.apply_discovered_rules_fast_mark_target(
        target_excel=target_excel,
        detect_rules_file=detect_rules_file,
        out_excel_fast=out_excel_fast,
        baseline_excel=baseline_excel,
        row_name=req.row_name,
        window_height=req.window_height,
        window_width=req.window_width,
        tolerance=req.tolerance,
        strict_row_match=req.strict_row_match,
        phase1_rules_file=phase1_rules_file,
    )

    return {
        "success": True,
        "mode": "fast_mark",
        "target_excel": target_excel,
        "detect_rules_file": detect_rules_file,
        "phase1_rules_file": phase1_rules_file,
        "out_excel_fast": out_path,
    }


def _run_rules_only(req: RulesOnlyRequest) -> Dict[str, Any]:
    target_excel = _abs_path(req.target_excel)
    detect_rules_file = _abs_path(req.detect_rules_file)
    out_excel = _abs_path(req.out_excel) if req.out_excel else str(_safe_outputs_dir() / "audit_marked.xlsx")
    out_excel = _ensure_parent_dir(out_excel)

    if not _exists(target_excel):
        raise _http_error(f"target_excel not found: {target_excel}")
    if not _exists(detect_rules_file):
        raise _http_error(f"detect_rules_file not found: {detect_rules_file}")

    out_path = ExcelStudio.detect_target_with_rules(
        target_excel=target_excel,
        detect_rules_file=detect_rules_file,
        out_excel=out_excel,
        row_name=req.row_name,
        window_height=req.window_height,
        window_width=req.window_width,
        tolerance=req.tolerance,
        strict_row_match=req.strict_row_match,
    )

    return {
        "success": True,
        "mode": "rules_only",
        "target_excel": target_excel,
        "detect_rules_file": detect_rules_file,
        "out_excel": out_path,
    }


def _run_full_flow(req: FullFlowRequest) -> Dict[str, Any]:
    baseline_excel = _abs_path(req.baseline_excel)
    target_excel = _abs_path(req.target_excel)
    detect_rules_file = _abs_path(req.detect_rules_file) if req.detect_rules_file else None
    out_excel = _abs_path(req.out_excel) if req.out_excel else str(_safe_outputs_dir() / "audit_marked.xlsx")
    out_excel = _ensure_parent_dir(out_excel)

    if not _exists(baseline_excel):
        raise _http_error(f"baseline_excel not found: {baseline_excel}")
    if not _exists(target_excel):
        raise _http_error(f"target_excel not found: {target_excel}")
    if detect_rules_file and not _exists(detect_rules_file):
        raise _http_error(f"detect_rules_file not found: {detect_rules_file}")

    parser = ExcelStudio.build_arg_parser()
    argv: List[str] = [baseline_excel, target_excel]

    if detect_rules_file:
        argv.extend(["--detect-rules-file", detect_rules_file])

    argv.extend(["--out-excel", out_excel])
    argv.extend(["--window-height", str(req.window_height)])
    argv.extend(["--window-width", str(req.window_width)])
    argv.extend(["--tolerance", str(req.tolerance)])

    if req.row_name:
        argv.extend(["--row-name", req.row_name])
    if req.strict_row_match:
        argv.append("--strict-row-match")
    if req.use_openai:
        argv.append("--use-openai")
    if req.openai_model:
        argv.extend(["--openai-model", req.openai_model])

    argv.extend(["--consistency-threshold", str(req.consistency_threshold)])
    argv.extend(["--quick-scan-threshold", str(req.quick_scan_threshold)])

    args = parser.parse_args(argv)
    ExcelStudio.main(args)

    auto_detect_rules = detect_rules_file or str(_safe_outputs_dir() / "discovered_rules.json")
    phase1_rules = str(_safe_outputs_dir() / "discovered_rules_phase1.json")
    fast_mark = str(_safe_outputs_dir() / "audit_marked_fast.xlsx")

    mark_summary: Optional[Dict[str, Any]] = None
    suspect_sheet = None
    suspect_cells: List[Dict[str, Any]] = []
    if Path(auto_detect_rules).exists():
        audit_parser = rule_audit_mark_excel.build_arg_parser()
        audit_argv: List[str] = [
            "--baseline-rules", auto_detect_rules,
            "--excel", target_excel,
            "--out-excel", out_excel,
            "--window-height", str(req.window_height),
            "--window-width", str(req.window_width),
            "--tolerance", str(req.tolerance),
        ]
        if req.row_name:
            audit_argv.extend(["--row-name", req.row_name])
        if req.strict_row_match:
            audit_argv.append("--strict-row-match")
        audit_args = audit_parser.parse_args(audit_argv)
        mark_ret: Dict[str, Any] = {}
        rule_audit_mark_excel.main(audit_args, ret=mark_ret)
        mark_summary = mark_ret
        suspect_sheet = mark_ret.get("suspect_sheet")
        suspect_cells = mark_ret.get("suspect_cells", [])

    return {
        "success": True,
        "mode": "full_flow",
        "baseline_excel": baseline_excel,
        "target_excel": target_excel,
        "request_argv": argv,
        "mark_summary": mark_summary,
        "suspect_sheet": suspect_sheet,
        "suspect_cells": suspect_cells,
        "outputs": {
            "detect_rules_file": auto_detect_rules if Path(auto_detect_rules).exists() else None,
            "phase1_rules_file": phase1_rules if Path(phase1_rules).exists() else None,
            "audit_marked_fast": fast_mark if Path(fast_mark).exists() else None,
            "audit_marked": out_excel if Path(out_excel).exists() else None,
            "rule_outputs": _existing_only(_collect_rule_output_paths()),
            "rule_outputs_phase1": _existing_only(_collect_rule_output_paths("_phase1")),
            "rule_outputs_phase2_final": _existing_only(_collect_rule_output_paths("_phase2_final")),
        },
    }


def _run_full_flow_json(req: FullFlowJsonRequest) -> Dict[str, Any]:
    # rule discovery on left table
    discovery_req = RuleDiscoveryJsonRequest(
        baseline_table=req.baseline_table,
        baseline_sheet_name=req.baseline_sheet_name,
        start_loc_row_name=req.row_name,
        window_height=req.window_height,
        window_width=req.window_width,
        use_openai=req.use_openai,
        openai_model=req.openai_model,
        consistency_threshold=req.consistency_threshold,
        quick_scan_threshold=req.quick_scan_threshold,
        quick_scan_seed=req.quick_scan_seed,
        use_phase1_global_llm_batch=req.use_phase1_global_llm_batch,
        phase2_overlap_phase1_retro=req.phase2_overlap_phase1_retro,
        step_rows=req.step_rows,
        step_cols=req.step_cols,
        degeneracy_min_nonzero_count=req.degeneracy_min_nonzero_count,
        degeneracy_min_distinct_nonzero=req.degeneracy_min_distinct_nonzero,
    )
    discovery = _run_rule_discovery_json(discovery_req)
    if not discovery.get("success"):
        return {
            "success": False,
            "mode": "full_flow_json",
            "error": "rule discovery failed",
            "discovery": discovery,
        }

    detect_rules_file = discovery.get("detect_rules_file") or str(_get_rules_file_path())
    if not _exists(detect_rules_file):
        raise _http_error(f"detect_rules_file not found: {detect_rules_file}")

    audit_req = AuditJsonRequest(
        detect_rules_file=detect_rules_file,
        target_table=req.target_table,
        target_sheet_name=req.target_sheet_name,
        row_name=req.row_name,
        window_height=req.window_height,
        window_width=req.window_width,
        tolerance=req.tolerance,
        strict_row_match=req.strict_row_match,
    )
    audit = _run_audit_json(audit_req)

    return {
        "success": True,
        "mode": "full_flow_json",
        "detect_rules_file": detect_rules_file,
        "discovery": discovery,
        "audit": audit,
        "suspect_sheet": audit.get("suspect_sheet"),
        "suspect_cells": audit.get("suspect_cells", []),
        "outputs": discovery.get("outputs", {}),
    }


def _run_rule_discovery_agent(req: RuleDiscoveryAgentRequest) -> Dict[str, Any]:
    baseline_table = _parse_table_text(req.baseline_text, req.table_format)
    discovery_req = RuleDiscoveryJsonRequest(
        baseline_table=baseline_table,
        baseline_sheet_name=req.baseline_sheet_name,
        start_loc_row_name=req.start_loc_row_name,
        window_height=req.window_height,
        window_width=req.window_width,
        use_openai=req.use_openai,
        openai_model=req.openai_model,
        consistency_threshold=req.consistency_threshold,
        quick_scan_threshold=req.quick_scan_threshold,
        quick_scan_seed=req.quick_scan_seed,
        use_phase1_global_llm_batch=req.use_phase1_global_llm_batch,
        phase2_overlap_phase1_retro=req.phase2_overlap_phase1_retro,
        step_rows=req.step_rows,
        step_cols=req.step_cols,
        degeneracy_min_nonzero_count=req.degeneracy_min_nonzero_count,
        degeneracy_min_distinct_nonzero=req.degeneracy_min_distinct_nonzero,
    )
    discovery = _run_rule_discovery_json(discovery_req)

    rules_payload = None
    detect_rules_file = discovery.get("detect_rules_file")
    if detect_rules_file and _exists(detect_rules_file):
        rules_payload = _read_json_file(Path(str(detect_rules_file)))

    rules_id = _store_rules_blob(rules_payload) if (rules_payload is not None and req.store_rules) else None
    rules_json = json.dumps(rules_payload, ensure_ascii=False) if (rules_payload is not None and req.return_rules_json) else None

    rows = len(baseline_table)
    cols = len(baseline_table[0]) if baseline_table else 0
    summary_text = (
        f"baseline_rows={rows}; baseline_cols={cols}; success={bool(discovery.get('success'))}; "
        f"rules_count={_estimate_rules_count(rules_payload)}"
    )

    return {
        "success": bool(discovery.get("success")),
        "mode": "rule_discovery_agent",
        "table_rows": rows,
        "table_cols": cols,
        "detect_rules_file": detect_rules_file,
        "rules_id": rules_id,
        "rules_count": _estimate_rules_count(rules_payload),
        "rules_json": rules_json,
        "summary_text": summary_text,
        "discovery": discovery,
    }


def _run_audit_agent(req: AuditAgentRequest) -> Dict[str, Any]:
    target_table = _parse_table_text(req.target_text, req.table_format)
    rules_pack = _resolve_rules_payload(
        rules_id=req.rules_id,
        rules_json=req.rules_json,
        rules=req.rules,
        detect_rules_file=req.detect_rules_file,
    )
    rules_payload = rules_pack["rules"]
    rules_source = str(rules_pack["source"])

    audit = _run_audit_json_with_rules(
        target_table=target_table,
        target_sheet_name=req.target_sheet_name,
        row_name=req.row_name,
        window_height=req.window_height,
        window_width=req.window_width,
        tolerance=req.tolerance,
        strict_row_match=req.strict_row_match,
        rules_payload=rules_payload,
        rules_source=rules_source,
    )

    rules_id = req.rules_id if rules_source == "rules_id" else None
    if rules_id is None and req.store_rules:
        rules_id = _store_rules_blob(rules_payload)
    rules_json = json.dumps(rules_payload, ensure_ascii=False) if req.return_rules_json else None

    return {
        "success": bool(audit.get("success")),
        "mode": "audit_agent",
        "table_rows": len(target_table),
        "table_cols": len(target_table[0]) if target_table else 0,
        "rules_source": rules_source,
        "rules_id": rules_id,
        "rules_count": _estimate_rules_count(rules_payload),
        "rules_json": rules_json,
        "summary_text": _build_suspect_summary_text(audit),
        "audit": audit,
        "suspect_sheet": audit.get("suspect_sheet"),
        "suspect_cells": audit.get("suspect_cells", []),
    }


def _run_full_flow_agent(req: FullFlowAgentRequest) -> Dict[str, Any]:
    baseline_table = _parse_table_text(req.baseline_text, req.baseline_table_format)
    target_table = _parse_table_text(req.target_text, req.target_table_format)
    has_input_rules = bool(req.rules_id or req.rules_json or req.rules is not None or req.detect_rules_file)

    discovery: Optional[Dict[str, Any]] = None
    rules_payload: Any = None
    rules_source = "discovery"
    detect_rules_file: Optional[str] = None

    if has_input_rules and not req.force_discovery:
        rules_pack = _resolve_rules_payload(
            rules_id=req.rules_id,
            rules_json=req.rules_json,
            rules=req.rules,
            detect_rules_file=req.detect_rules_file,
        )
        rules_payload = rules_pack["rules"]
        rules_source = str(rules_pack["source"])
    else:
        discovery_req = RuleDiscoveryJsonRequest(
            baseline_table=baseline_table,
            baseline_sheet_name=req.baseline_sheet_name,
            start_loc_row_name=req.start_loc_row_name,
            window_height=req.window_height,
            window_width=req.window_width,
            use_openai=req.use_openai,
            openai_model=req.openai_model,
            consistency_threshold=req.consistency_threshold,
            quick_scan_threshold=req.quick_scan_threshold,
            quick_scan_seed=req.quick_scan_seed,
            use_phase1_global_llm_batch=req.use_phase1_global_llm_batch,
            phase2_overlap_phase1_retro=req.phase2_overlap_phase1_retro,
            step_rows=req.step_rows,
            step_cols=req.step_cols,
            degeneracy_min_nonzero_count=req.degeneracy_min_nonzero_count,
            degeneracy_min_distinct_nonzero=req.degeneracy_min_distinct_nonzero,
        )
        discovery = _run_rule_discovery_json(discovery_req)
        if not discovery.get("success"):
            return {
                "success": False,
                "mode": "full_flow_agent",
                "error": "rule discovery failed",
                "discovery": discovery,
            }
        detect_rules_file = discovery.get("detect_rules_file")
        if not detect_rules_file or not _exists(detect_rules_file):
            raise _http_error(f"detect_rules_file not found: {detect_rules_file}")
        rules_payload = _read_json_file(Path(str(detect_rules_file)))

    audit = _run_audit_json_with_rules(
        target_table=target_table,
        target_sheet_name=req.target_sheet_name,
        row_name=req.row_name,
        window_height=req.window_height,
        window_width=req.window_width,
        tolerance=req.tolerance,
        strict_row_match=req.strict_row_match,
        rules_payload=rules_payload,
        rules_source=rules_source,
    )

    rules_id = req.rules_id if rules_source == "rules_id" else None
    if rules_id is None and req.store_rules:
        rules_id = _store_rules_blob(rules_payload)
    rules_json = json.dumps(rules_payload, ensure_ascii=False) if req.return_rules_json else None

    return {
        "success": True,
        "mode": "full_flow_agent",
        "flow_branch": "rules_input" if (has_input_rules and not req.force_discovery) else "discover_then_audit",
        "rules_source": rules_source,
        "detect_rules_file": detect_rules_file,
        "rules_id": rules_id,
        "rules_count": _estimate_rules_count(rules_payload),
        "rules_json": rules_json,
        "summary_text": _build_suspect_summary_text(audit),
        "baseline_table_rows": len(baseline_table),
        "target_table_rows": len(target_table),
        "discovery": discovery,
        "audit": audit,
        "suspect_sheet": audit.get("suspect_sheet"),
        "suspect_cells": audit.get("suspect_cells", []),
    }


# ---------------------------------------------------------
# Task runners
# ---------------------------------------------------------
def _ensure_not_cancelled(cancel_event: threading.Event) -> None:
    if cancel_event.is_set():
        raise _TaskCancelled()


def _run_task_action(action: str, payload: Dict[str, Any], cancel_event: threading.Event) -> Dict[str, Any]:
    _ensure_not_cancelled(cancel_event)
    if action == "ruleDiscovery":
        _set_task_progress(15, "規則偵測中")
        req = RuleDiscoveryJsonRequest(**payload)
        result = _run_rule_discovery_json(req)
        _set_task_progress(90, "規則偵測完成")
        return result

    if action in {"audit", "markFast", "rulesOnly"}:
        _set_task_progress(20, "稽核中")
        req = AuditJsonRequest(**payload)
        result = _run_audit_json(req)
        _set_task_progress(90, "稽核完成")
        return result

    if action == "fullFlow":
        req = FullFlowJsonRequest(**payload)
        _set_task_progress(15, "規則偵測中")
        discovery_req = RuleDiscoveryJsonRequest(
            baseline_table=req.baseline_table,
            baseline_sheet_name=req.baseline_sheet_name,
            start_loc_row_name=req.row_name,
            window_height=req.window_height,
            window_width=req.window_width,
            use_openai=req.use_openai,
            openai_model=req.openai_model,
            consistency_threshold=req.consistency_threshold,
            quick_scan_threshold=req.quick_scan_threshold,
            quick_scan_seed=req.quick_scan_seed,
            use_phase1_global_llm_batch=req.use_phase1_global_llm_batch,
            phase2_overlap_phase1_retro=req.phase2_overlap_phase1_retro,
            step_rows=req.step_rows,
            step_cols=req.step_cols,
            degeneracy_min_nonzero_count=req.degeneracy_min_nonzero_count,
            degeneracy_min_distinct_nonzero=req.degeneracy_min_distinct_nonzero,
        )
        discovery = _run_rule_discovery_json(discovery_req)
        if not discovery.get("success"):
            return {
                "success": False,
                "mode": "full_flow_json",
                "error": "rule discovery failed",
                "discovery": discovery,
            }
        _ensure_not_cancelled(cancel_event)

        _set_task_progress(65, "稽核中")
        detect_rules_file = discovery.get("detect_rules_file") or str(_get_rules_file_path())
        if not _exists(detect_rules_file):
            raise _http_error(f"detect_rules_file not found: {detect_rules_file}")
        audit_req = AuditJsonRequest(
            detect_rules_file=detect_rules_file,
            target_table=req.target_table,
            target_sheet_name=req.target_sheet_name,
            row_name=req.row_name,
            window_height=req.window_height,
            window_width=req.window_width,
            tolerance=req.tolerance,
            strict_row_match=req.strict_row_match,
        )
        audit = _run_audit_json(audit_req)
        _set_task_progress(90, "稽核完成")
        return {
            "success": True,
            "mode": "full_flow_json",
            "detect_rules_file": detect_rules_file,
            "discovery": discovery,
            "audit": audit,
            "suspect_sheet": audit.get("suspect_sheet"),
            "suspect_cells": audit.get("suspect_cells", []),
            "outputs": discovery.get("outputs", {}),
        }

    raise _http_error(f"unsupported task action: {action}")


def _task_worker(task_id: str, action: str, payload: Dict[str, Any], cancel_event: threading.Event) -> None:
    try:
        _set_task_state(
            task_id=task_id,
            status="running",
            action=action,
            progress=5,
            message="準備中",
            result=None,
            error=None,
            started_at=_utc_now(),
            cancel_event=cancel_event,
        )
        result = _run_task_action(action, payload, cancel_event)
        if cancel_event.is_set():
            raise _TaskCancelled()
        _set_task_state(status="done", progress=100, message="完成", result=result)
    except _TaskCancelled:
        _set_task_state(status="cancelled", progress=0, message="已中止", result=None)
    except Exception as exc:
        _set_task_state(status="error", progress=100, message=str(exc), error=traceback.format_exc())


def _start_task(action: str, payload: Dict[str, Any]) -> str:
    with TASK_LOCK:
        if TASK_STATE.get("status") in {"running", "cancel_requested"}:
            raise _http_error("task already running", status_code=409)
        if action not in {"ruleDiscovery", "audit", "markFast", "rulesOnly", "fullFlow"}:
            raise _http_error(f"unsupported task action: {action}")
        task_id = str(uuid.uuid4())
        cancel_event = threading.Event()
        TASK_STATE.update(
            {
                "task_id": task_id,
                "status": "running",
                "action": action,
                "progress": 0,
                "message": "排程中",
                "result": None,
                "error": None,
                "started_at": _utc_now(),
                "updated_at": _utc_now(),
                "cancel_event": cancel_event,
            }
        )
    thread = threading.Thread(
        target=_task_worker,
        args=(task_id, action, payload, cancel_event),
        daemon=True,
    )
    thread.start()
    return task_id


# ---------------------------------------------------------
# API routes
# ---------------------------------------------------------
@app.get("/")
def root():
    return {
        "service": "ExcelStudio API Server",
        "version": "1.0.0",
        "endpoints": [
            "GET /health",
            "POST /api/task/start",
            "GET /api/task/progress",
            "POST /api/task/stop",
            "GET /api/rules/discovered",
            "POST /api/excel/upload",
            "POST /api/outputs/clear",
            "POST /api/rules/discover",
            "POST /api/rules/discover-json",
            "POST /api/audit",
            "POST /api/audit-json",
            "POST /api/rules/discover-agent",
            "POST /api/audit-agent",
            "POST /api/full-flow-agent",
            "POST /api/mark-fast",
            "POST /api/rules-only",
            "POST /api/full-flow",
            "POST /api/full-flow-json",
            "POST /api/xlsx/command",
            "POST /api/xlsx/preview-group-aggregate",
            "POST /api/xlsx/group-aggregate",
            "POST /api/xlsx/capture-range-image",
            "POST /api/xlsx/capture-range-figure",
        ],
    }


@app.get("/health")
def health():
    outputs_dir = _safe_outputs_dir()
    return {
        "ok": True,
        "cwd": str(CURRENT_DIR),
        "outputs_dir": str(outputs_dir),
        "excelstudio_py": str(CURRENT_DIR / "ExcelStudio.py"),
        "data_process_dir": str(DATA_PROCESS_DIR),
    }


@app.post("/api/task/start")
async def api_task_start(req: TaskStartRequest):
    task_id = _start_task(req.action, req.payload or {})
    return {"success": True, "task_id": task_id}


@app.get("/api/task/progress")
def api_task_progress(task_id: Optional[str] = None):
    state = _task_public_state()
    if task_id and state.get("task_id") != task_id:
        raise _http_error("task_id not found", status_code=404)
    return {"success": True, **state}


@app.post("/api/task/stop")
async def api_task_stop(req: TaskStopRequest):
    with TASK_LOCK:
        if TASK_STATE.get("status") != "running":
            return {
                "success": True,
                "status": TASK_STATE.get("status"),
                "task_id": TASK_STATE.get("task_id"),
            }
        if req.task_id and req.task_id != TASK_STATE.get("task_id"):
            raise _http_error("task_id not found", status_code=404)
        cancel_event = TASK_STATE.get("cancel_event")
        if cancel_event:
            cancel_event.set()
        TASK_STATE["status"] = "cancel_requested"
        TASK_STATE["message"] = "中止中"
        TASK_STATE["updated_at"] = _utc_now()
    return {"success": True, "status": "cancel_requested", "task_id": TASK_STATE.get("task_id")}


@app.get("/api/rules/discovered")
def api_rules_discovered():
    path = _get_rules_file_path()
    data = _read_json_file(path)
    return {
        "success": True,
        "file": str(path),
        "data": data,
    }


@app.get("/api/rules/path")
def api_rules_path():
    path = _get_rules_file_path()
    return {"success": True, "path": str(path)}


@app.get("/api/rules/download")
def api_rules_download():
    path = _get_rules_file_path()
    if not path.exists():
        raise _http_error(f"rules file not found: {path}")
    return FileResponse(
        str(path),
        media_type="application/json",
        filename=path.name,
    )


@app.post("/api/rules/upload")
async def api_rules_upload(file: UploadFile = File(...)):
    if not file.filename:
        raise _http_error("missing file")
    if not file.filename.lower().endswith(".json"):
        raise _http_error("rules file must be .json")
    content = await file.read()
    try:
        json.loads(content.decode("utf-8-sig"))
    except Exception:
        raise _http_error("invalid json file")

    path = _get_rules_file_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "wb") as f:
        f.write(content)
    return {"success": True, "path": str(path), "size": len(content)}


@app.post("/api/excel/upload")
async def api_excel_upload(file: UploadFile = File(...)):
    if not file.filename:
        raise _http_error("missing file")
    filename = _safe_filename(file.filename)
    ext = Path(filename).suffix.lower()
    if ext not in {".xlsx", ".xls", ".xlsm", ".csv"}:
        raise _http_error("excel file must be .xlsx/.xls/.xlsm/.csv")

    content = await file.read()
    uploads_dir = _safe_uploads_dir()
    target = uploads_dir / filename
    with open(target, "wb") as f:
        f.write(content)
    return {"success": True, "path": str(target), "size": len(content)}


@app.post("/api/outputs/clear")
def api_outputs_clear():
    result = _clear_outputs_dir()
    return {"success": True, **result}


@app.post("/api/rules/discover")
async def api_rules_discover(req: RuleDiscoveryRequest):
    try:
        return await run_in_threadpool(_run_rule_discovery, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "rule discovery failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/rules/discover-json")
async def api_rules_discover_json(req: RuleDiscoveryJsonRequest):
    try:
        return await run_in_threadpool(_run_rule_discovery_json, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "rule discovery (json) failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/audit")
async def api_audit(req: AuditRequest):
    try:
        return await run_in_threadpool(_run_audit, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "audit failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/audit-json")
async def api_audit_json(req: AuditJsonRequest):
    try:
        return await run_in_threadpool(_run_audit_json, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "audit (json) failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/mark-fast")
async def api_mark_fast(req: FastMarkRequest):
    try:
        return await run_in_threadpool(_run_fast_mark, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "fast mark failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/rules-only")
async def api_rules_only(req: RulesOnlyRequest):
    try:
        return await run_in_threadpool(_run_rules_only, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "rules-only audit failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/full-flow")
async def api_full_flow(req: FullFlowRequest):
    try:
        return await run_in_threadpool(_run_full_flow, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "full flow failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/full-flow-json")
async def api_full_flow_json(req: FullFlowJsonRequest):
    try:
        return await run_in_threadpool(_run_full_flow_json, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "full flow (json) failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/rules/discover-agent")
async def api_rules_discover_agent(req: RuleDiscoveryAgentRequest):
    try:
        return await run_in_threadpool(_run_rule_discovery_agent, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "rule discovery (agent) failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/audit-agent")
async def api_audit_agent(req: AuditAgentRequest):
    try:
        return await run_in_threadpool(_run_audit_agent, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "audit (agent) failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/full-flow-agent")
async def api_full_flow_agent(req: FullFlowAgentRequest):
    try:
        return await run_in_threadpool(_run_full_flow_agent, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "full flow (agent) failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/xlsx/command")
async def api_xlsx_command(req: XlsxStdioCommandRequest):
    try:
        return await run_in_threadpool(_run_xlsx_stdio_command, req)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "xlsx stdio command failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/xlsx/preview-group-aggregate")
async def api_xlsx_preview_group_aggregate(req: PreviewGroupAggregateRequest):
    try:
        args = req.model_dump(exclude_none=True)
        request_id = args.pop("request_id", None)
        payload = XlsxStdioCommandRequest(
            command="preview_group_aggregate",
            args=args,
            request_id=request_id,
        )
        return await run_in_threadpool(_run_xlsx_stdio_command, payload)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "preview_group_aggregate failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/xlsx/group-aggregate")
async def api_xlsx_group_aggregate(req: GroupAggregateRequest):
    try:
        args = req.model_dump(exclude_none=True)
        request_id = args.pop("request_id", None)

        # 對齊 xlsx_stdio.py 既有參數名
        if "start_cell_out" in args:
            args["start_cell"] = args.pop("start_cell_out")

        payload = XlsxStdioCommandRequest(
            command="group_aggregate",
            args=args,
            request_id=request_id,
        )
        return await run_in_threadpool(_run_xlsx_stdio_command, payload)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "group_aggregate failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/xlsx/capture-range-image")
async def api_xlsx_capture_range_image(req: CaptureRangeImageRequest):
    try:
        args = req.model_dump(exclude_none=True)
        request_id = args.pop("request_id", None)
        payload = XlsxStdioCommandRequest(
            command="capture_range_image",
            args=args,
            request_id=request_id,
        )
        return await run_in_threadpool(_run_xlsx_stdio_command, payload)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "capture_range_image failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )


@app.post("/api/xlsx/capture-range-figure")
async def api_xlsx_capture_range_figure(req: CaptureRangeFigureRequest):
    try:
        args = req.model_dump(exclude_none=True)
        request_id = args.pop("request_id", None)
        payload = XlsxStdioCommandRequest(
            command="capture_range_figure",
            args=args,
            request_id=request_id,
        )
        return await run_in_threadpool(_run_xlsx_stdio_command, payload)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail={
                "message": "capture_range_figure failed",
                "error": str(e),
                "traceback": traceback.format_exc(),
            },
        )
# ---------------------------------------------------------
# Local run
# ---------------------------------------------------------
if __name__ == "__main__":
    import uvicorn

    cfg_path = CURRENT_DIR / "config.json"
    try:
        cfg_preview = json.dumps(config, ensure_ascii=False)
    except Exception:
        cfg_preview = str(config)

    host = os.getenv(
        "EXCELSTUDIO_API_HOST",
        str(_get_config_value(config, "api_host", "host", default="0.0.0.0")),
    )
    port = int(
        os.getenv(
            "EXCELSTUDIO_API_PORT",
            str(_get_config_value(config, "api_port", "port", default="7018")),
        )
    )

    print(
        "[BOOT] ExcelStudio api_server.py",
        f"config={cfg_path if cfg_path.exists() else 'missing'}",
        f"loaded={cfg_preview}",
        f"host={host}",
        f"port={port}",
    )

    uvicorn.run(
        "api_server:app",
        host=host,
        port=port,
        reload=False,
    )
