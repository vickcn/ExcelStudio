#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
XLSX 解析器 - 符合 ContextParser 統一架構

此模組提供 Excel / Spreadsheet 檔案解析功能，遵循 parser_design.md 定義的核心功能：
- preview: 使用 pandas 將每個 worksheet 輸出為 Markdown 表格文字
- extract: 先以 sheet 為單位建立 unit_paras，並蒐集公式 / 圖像等 metadata
- segment: 在每個 sheet 中偵測實際 table 區塊，必要時批次做圖文解釋
- chunk: 依 row / column / cell 模式，從每個 table segment 展開最終 chunk
- process: extract -> segment -> chunk 一條龍流程
"""

from __future__ import annotations

import os
import sys
import io
import re
import csv
import json
import math
import base64
import shutil
import logging
import subprocess
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass
from openpyxl.utils import get_column_letter
from copy import deepcopy as dcp

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    pd = None
    PANDAS_AVAILABLE = False

try:
    import numpy as np
    NUMPY_AVAILABLE = True
except ImportError:
    np = None
    NUMPY_AVAILABLE = False

try:
    from openpyxl import load_workbook, Workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    Workbook = None
    load_workbook = None
    OPENPYXL_AVAILABLE = False

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

try:
    from ContextParser.context_parser import parse_keywords_from_text
    from ContextParser.context_parser import _merge_heading_tags_by_orders
    from ContextParser.context_parser import dedup_multi_prompts_by_llm
    from ContextParser.context_parser import filter_meaningless_tags
    from ContextParser.context_parser import analyze_images_via_batch_common
except Exception:
    try:
        from .context_parser import parse_keywords_from_text
        from .context_parser import _merge_heading_tags_by_orders
        from .context_parser import dedup_multi_prompts_by_llm
        from .context_parser import filter_meaningless_tags
        from .context_parser import analyze_images_via_batch_common
    except Exception:
        parse_keywords_from_text = None
        _merge_heading_tags_by_orders = None
        dedup_multi_prompts_by_llm = None
        filter_meaningless_tags = None
        analyze_images_via_batch_common = None

DATA_PROCESS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if DATA_PROCESS_DIR not in sys.path:
    sys.path.insert(0, DATA_PROCESS_DIR)
try:
    from package import LOGger
    from package import dataframeprocedure as DFP
except Exception:
    try:
        import LOGger  # type: ignore
        import dataframeprocedure as DFP  # type: ignore
    except Exception:
        LOGger = None
        DFP = None

m_fn = os.path.basename(__file__).replace('.py', '')


def _setup_logger() -> logging.Logger:
    if LOGger is not None:
        log_dir = os.path.join(os.path.dirname(__file__), 'log')
        os.makedirs(log_dir, exist_ok=True)
        logger = LOGger.addloger(logfile=os.path.join(log_dir, 'context_parser_%t.log'))
        logger.error = lambda x, *args, colora=getattr(LOGger, 'FAIL', None), **kwargs: logger(x, *args, **kwargs, colora=colora)
        logger.warning = lambda x, *args, colora=getattr(LOGger, 'WARNING', None), **kwargs: logger(x, *args, **kwargs, colora=colora)
        logger.info = lambda x, *args, **kwargs: logger(x, *args, **kwargs)
        logger.debug = lambda x, *args, colora=getattr(LOGger, 'OKBLUE', None), **kwargs: logger(x, *args, **kwargs, colora=colora)
        logger.exception = lambda x, logfile='', **kwargs: LOGger.exception_process(x, logfile=logfile, **kwargs)
        logger.summary = lambda x, *args, colora=getattr(LOGger, 'OKCYAN', None), **kwargs: logger(x, *args, **kwargs, colora=colora)
        return logger

    logger = logging.getLogger(m_fn)
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
        logger.addHandler(handler)
    return logger


m_logger = _setup_logger()
m_config_file = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'config.json')
m_config = {}
if LOGger is not None and os.path.exists(m_config_file):
    try:
        m_config = LOGger.load_json(m_config_file)
    except Exception:
        m_config = {}


def _log_prefix(color_name: str) -> str:
    if LOGger is None:
        return ''
    return getattr(LOGger, color_name, '') or ''


# ============================================================
# 輔助函式
# ============================================================

# dataclass reserved for next refactor; current pipeline still uses dict for backward compatibility
@dataclass
class HeaderInfo:
    has_header_row: bool
    has_header_col: bool
    header_rows: list[int]
    header_cols: list[int]
    headers: list[str]
    confidence: float


@dataclass
class IndexInfo:
    has_index_col: bool
    has_index_row: bool
    index_cols: list[int]
    index_rows: list[int]
    index_values: list[str]
    confidence: float


@dataclass
class TableMetadata:
    memo: list[str]
    formulas: list[dict]
    images: list[dict]
    merged_cells: list[dict]
    original_bbox: tuple[int, int, int, int]
    sheet_bbox: tuple[int, int, int, int] | None


@dataclass
class TableRegion:
    table_id: int
    sheet_name: str
    sheet_index: int

    bbox: tuple[int, int, int, int]
    shape: tuple[int, int]
    non_empty_count: int
    density: float

    table_df: Any
    values: list[list[Any]]
    markdown: str

    header: HeaderInfo
    index: IndexInfo
    metadata: TableMetadata

    region_type: str = "unknown"
    confidence: float = 0.0
    reason: str = ""


def _require_dependencies() -> None:
    if not PANDAS_AVAILABLE:
        raise ImportError('pandas 未安裝，請執行: pip install pandas')
    if not OPENPYXL_AVAILABLE:
        raise ImportError('openpyxl 未安裝，請執行: pip install openpyxl')


class _MiniDFP:
    @staticmethod
    def rpathrpt(path: str) -> str:
        p = Path(path)
        stem = p.stem
        suffix = p.suffix
        parent = p.parent
        idx = 1
        candidate = p
        while candidate.exists():
            candidate = parent / f'{stem}_{idx}{suffix}'
            idx += 1
        return str(candidate)


def _rpathrpt(path: str) -> str:
    if DFP is not None and hasattr(DFP, 'rpathrpt'):
        return DFP.rpathrpt(path)
    return _MiniDFP.rpathrpt(path)



def _normalize_none_like(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, str) and value.strip() == '':
        return None
    return value



def _truncate_text(text: str, max_len: Optional[int] = None) -> str:
    if max_len is None or max_len <= 0:
        return text
    if len(text) <= max_len:
        return text
    return text[:max_len] + '...'


def _cell_to_display_text(value: Any, max_len: Optional[int] = None, digit: int = 6) -> str:
    value = _normalize_none_like(value)
    if value is None:
        return ''
    if isinstance(value, (int, float)) or (NUMPY_AVAILABLE and isinstance(value, np.generic)):
        if DFP is not None and hasattr(DFP, 'parse'):
            try:
                return _truncate_text(str(DFP.parse(value, digit=digit)), max_len=max_len)
            except Exception:
                return _truncate_text(str(value), max_len=max_len)
    if isinstance(value, bytes):
        try:
            return _truncate_text(value.decode('utf-8', errors='ignore'), max_len=max_len)
        except Exception:
            return _truncate_text(str(value), max_len=max_len)
    return _truncate_text(str(value), max_len=max_len)



def _safe_sheet_title(name: str) -> str:
    return re.sub(r'[\\/*?:\[\]]+', '_', name or 'Sheet')[:31] or 'Sheet'



def _read_text_with_fallbacks(file_path: str, encodings: Optional[List[str]] = None) -> str:
    encodings = encodings or ['utf-8-sig', 'utf-8', 'cp950', 'big5', 'latin1']
    last_error = None
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc, newline='') as f:
                return f.read()
        except Exception as e:
            last_error = e
    raise ValueError(f'無法讀取 CSV 文字內容: {last_error}')



def _detect_csv_dialect(text: str) -> csv.Dialect:
    sample = text[:8192]
    try:
        return csv.Sniffer().sniff(sample, delimiters=',;\t|')
    except Exception:
        class _Default(csv.Dialect):
            delimiter = ','
            quotechar = '"'
            doublequote = True
            skipinitialspace = False
            lineterminator = '\n'
            quoting = csv.QUOTE_MINIMAL
        return _Default



def _to_serializable_df(df: 'pd.DataFrame') -> List[List[Any]]:
    safe_df = df.where(pd.notna(df), None)
    rows: List[List[Any]] = []
    for row in safe_df.values.tolist():
        rows.append([_normalize_none_like(v) for v in row])
    return rows



def _df_to_markdown(df: 'pd.DataFrame', max_rows: Optional[int] = None, max_cols: Optional[int] = None) -> str:
    if df is None:
        return ''
    view = df.copy()
    if max_rows is not None:
        view = view.head(max_rows)
    if max_cols is not None:
        view = view.iloc[:, :max_cols]
    try:
        return view.fillna('').to_markdown(index=False)
    except Exception:
        return view.fillna('').to_csv(index=False)



def _normalize_to_xlsx(file_path: str, output_dir: Optional[str] = None) -> str:
    _require_dependencies()
    if not os.path.exists(file_path):
        raise FileNotFoundError(f'檔案不存在: {file_path}')

    src = Path(file_path)
    suffix = src.suffix.lower()
    if suffix == '.xlsx':
        return str(src)
    if suffix == '.csv':
        return _convert_csv_to_xlsx(file_path, output_dir=output_dir)
    if suffix == '.xlsm':
        return _convert_xlsm_to_xlsx(file_path, output_dir=output_dir)
    if suffix == '.xls':
        return _convert_xls_to_xlsx(file_path, output_dir=output_dir)
    raise ValueError(f'不支援的檔案格式: {file_path}。僅支援 .csv / .xls / .xlsm / .xlsx')



def _prepare_output_xlsx_path(file_path: str, output_dir: Optional[str] = None) -> str:
    src = Path(file_path)
    parent = Path(output_dir) if output_dir else src.parent
    parent.mkdir(parents=True, exist_ok=True)
    target = parent / f'{src.stem}.xlsx'
    if str(target) == str(src):
        target = parent / f'{src.stem}_normalized.xlsx'
    if target.exists():
        src_mtime = os.path.getmtime(file_path)
        try:
            if os.path.getmtime(target) >= src_mtime:
                return str(target)
        except Exception:
            pass
        return _rpathrpt(str(target))
    return str(target)


def _convert_csv_to_xlsx(file_path: str, output_dir: Optional[str] = None) -> str:
    out_path = _prepare_output_xlsx_path(file_path, output_dir=output_dir)
    raw_text = _read_text_with_fallbacks(file_path)
    dialect = _detect_csv_dialect(raw_text)
    df = pd.read_csv(io.StringIO(raw_text), sep=dialect.delimiter, engine='python')
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    m_logger.info(f'[_convert_csv_to_xlsx] 轉換成功: {file_path} -> {out_path}')
    return out_path



def _convert_xlsm_to_xlsx(file_path: str, output_dir: Optional[str] = None) -> str:
    out_path = _prepare_output_xlsx_path(file_path, output_dir=output_dir)
    wb = load_workbook(file_path, data_only=False)
    out_wb = Workbook()
    if out_wb.active:
        out_wb.remove(out_wb.active)
    for ws in wb.worksheets:
        new_ws = out_wb.create_sheet(title=_safe_sheet_title(ws.title))
        for row in ws.iter_rows():
            for cell in row:
                new_ws[cell.coordinate].value = cell.value
        for merged in ws.merged_cells.ranges:
            try:
                new_ws.merge_cells(str(merged))
            except Exception:
                pass
    out_wb.save(out_path)
    m_logger.info(f'[_convert_xlsm_to_xlsx] 轉換成功: {file_path} -> {out_path}')
    return out_path



def _convert_xls_to_xlsx(file_path: str, output_dir: Optional[str] = None) -> str:
    out_path = _prepare_output_xlsx_path(file_path, output_dir=output_dir)
    src_path = Path(file_path)

    libreoffice_exe = shutil.which('libreoffice') or shutil.which('soffice')
    if not libreoffice_exe and m_config:
        libreoffice_exe = m_config.get('doc_conversion', {}).get('libreoffice_path', '') or libreoffice_exe
    if libreoffice_exe:
        cmd = [libreoffice_exe, '--headless', '--convert-to', 'xlsx', '--outdir', str(Path(out_path).parent), str(src_path)]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            generated = Path(out_path).parent / f'{src_path.stem}.xlsx'
            if result.returncode == 0 and generated.exists():
                if str(generated) != out_path:
                    shutil.move(str(generated), out_path)
                m_logger.info(f'[_convert_xls_to_xlsx] LibreOffice 轉換成功: {file_path} -> {out_path}')
                return out_path
            m_logger.warning(f'[_convert_xls_to_xlsx] LibreOffice 轉換失敗: {result.stderr}')
        except Exception as e:
            m_logger.warning(f'[_convert_xls_to_xlsx] LibreOffice 轉換失敗: {e}')

    try:
        import xlrd  # type: ignore
        book = xlrd.open_workbook(file_path)
        out_wb = Workbook()
        if out_wb.active:
            out_wb.remove(out_wb.active)
        for sheet in book.sheets():
            ws = out_wb.create_sheet(title=_safe_sheet_title(sheet.name))
            for r in range(sheet.nrows):
                for c in range(sheet.ncols):
                    ws.cell(row=r + 1, column=c + 1, value=sheet.cell_value(r, c))
        out_wb.save(out_path)
        m_logger.info(f'[_convert_xls_to_xlsx] xlrd 轉換成功: {file_path} -> {out_path}')
        return out_path
    except Exception as e:
        raise ValueError(
            f'無法轉換 XLS 檔案 {file_path}: {e}。請安裝 LibreOffice 或 xlrd。'
        )



def _worksheet_to_dataframe(ws, include_formulas: bool = True) -> 'pd.DataFrame':
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0
    data: List[List[Any]] = []
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        row_values = []
        for cell in row:
            value = cell.value
            if hasattr(cell, 'data_type') and cell.data_type == 'f' and include_formulas:
                value = value if value is not None else ''
            row_values.append(value)
        data.append(row_values)
    if not data:
        return pd.DataFrame()
    return pd.DataFrame(data)



def _trim_empty_df(df: 'pd.DataFrame') -> Tuple['pd.DataFrame', Optional[Tuple[int, int, int, int]]]:
    if df is None or df.empty:
        return pd.DataFrame(), None
    mask = df.applymap(lambda x: _normalize_none_like(x) is not None)
    if not mask.to_numpy().any():
        return pd.DataFrame(), None
    non_empty_rows = np.where(mask.any(axis=1).to_numpy())[0]
    non_empty_cols = np.where(mask.any(axis=0).to_numpy())[0]
    top = int(non_empty_rows[0])
    bottom = int(non_empty_rows[-1])
    left = int(non_empty_cols[0])
    right = int(non_empty_cols[-1])
    trimmed = df.iloc[top:bottom + 1, left:right + 1].copy()
    return trimmed, (top + 1, left + 1, bottom + 1, right + 1)



def _extract_sheet_images(ws, include_images: bool = True, to_base64: bool = True) -> List[Dict[str, Any]]:
    results: List[Dict[str, Any]] = []
    images = list(getattr(ws, '_images', []) or [])
    for idx, img in enumerate(images):
        info: Dict[str, Any] = {
            'local_index': idx,
            'sheet_name': ws.title,
        }
        image_id = None
        anchor = getattr(img, 'anchor', None)
        anchor_cell = None
        top_row = None
        left_col = None
        if anchor is not None:
            _from = getattr(anchor, '_from', None)
            if _from is not None:
                top_row = getattr(_from, 'row', None)
                left_col = getattr(_from, 'col', None)
                if top_row is not None and left_col is not None:
                    anchor_cell = f'{_colnum_to_letter(left_col + 1)}{top_row + 1}'
        info['anchor_cell'] = anchor_cell
        info['anchor_row'] = (top_row + 1) if top_row is not None else None
        info['anchor_col'] = (left_col + 1) if left_col is not None else None
        info['type'] = 'cell_image' if anchor_cell else 'sheet_image'

        placeholder = None
        if include_images:
            image_id = idx
            placeholder = _make_image_placeholder(image_id)
            info['id'] = image_id
            info['placeholder'] = placeholder

        raw_bytes = None
        try:
            if hasattr(img, '_data'):
                raw_bytes = img._data()
            elif hasattr(img, 'ref') and isinstance(img.ref, (bytes, bytearray)):
                raw_bytes = bytes(img.ref)
        except Exception:
            raw_bytes = None

        if raw_bytes is not None:
            info['size_bytes'] = len(raw_bytes)
            if to_base64:
                info['base64'] = base64.b64encode(raw_bytes).decode('utf-8')
            else:
                info['bytes'] = raw_bytes
        results.append(info)
    return results



def _extract_formula_cells(ws) -> List[Dict[str, Any]]:
    formulas: List[Dict[str, Any]] = []
    for row in ws.iter_rows():
        for cell in row:
            try:
                if cell.data_type == 'f' or (isinstance(cell.value, str) and cell.value.startswith('=')):
                    formulas.append({
                        'sheet_name': ws.title,
                        'cell': cell.coordinate,
                        'formula': cell.value,
                    })
            except Exception:
                continue
    return formulas



def _colnum_to_letter(col_num: int) -> str:
    result = ''
    while col_num > 0:
        col_num, rem = divmod(col_num - 1, 26)
        result = chr(65 + rem) + result
    return result



def _make_image_placeholder(image_id: int) -> str:
    return f'[IMAGE_PLACEHOLDER_{image_id}]'



def _find_connected_components(mask: 'np.ndarray') -> List[Tuple[int, int, int, int]]:
    rows, cols = mask.shape
    visited = np.zeros_like(mask, dtype=bool)
    components: List[Tuple[int, int, int, int]] = []
    directions = [(-1, 0), (1, 0), (0, -1), (0, 1)]

    for r in range(rows):
        for c in range(cols):
            if not mask[r, c] or visited[r, c]:
                continue
            stack = [(r, c)]
            visited[r, c] = True
            min_r = max_r = r
            min_c = max_c = c
            while stack:
                cr, cc = stack.pop()
                min_r = min(min_r, cr)
                max_r = max(max_r, cr)
                min_c = min(min_c, cc)
                max_c = max(max_c, cc)
                for dr, dc in directions:
                    nr, nc = cr + dr, cc + dc
                    if 0 <= nr < rows and 0 <= nc < cols and mask[nr, nc] and not visited[nr, nc]:
                        visited[nr, nc] = True
                        stack.append((nr, nc))
            components.append((min_r, min_c, max_r, max_c))
    return components



def _detect_tables_from_sheet_df(sheet_df: 'pd.DataFrame', min_non_empty_cells: int = 1) -> List[Dict[str, Any]]:
    if sheet_df is None or sheet_df.empty:
        return []
    mask_df = sheet_df.applymap(lambda x: _normalize_none_like(x) is not None)
    mask = mask_df.to_numpy(dtype=bool)
    if not mask.any():
        return []

    boxes = _find_connected_components(mask)
    tables: List[Dict[str, Any]] = []
    for idx, (min_r, min_c, max_r, max_c) in enumerate(boxes):
        sub_df = sheet_df.iloc[min_r:max_r + 1, min_c:max_c + 1].copy()
        non_empty_count = int(sub_df.applymap(lambda x: _normalize_none_like(x) is not None).to_numpy().sum())
        if non_empty_count < min_non_empty_cells:
            continue
        tables.append({
            'table_id': idx,
            'top_row': min_r + 1,
            'left_col': min_c + 1,
            'bottom_row': max_r + 1,
            'right_col': max_c + 1,
            'shape': [int(sub_df.shape[0]), int(sub_df.shape[1])],
            'table_df': sub_df,
            'table_values': _to_serializable_df(sub_df),
            'table_markdown': _df_to_markdown(sub_df),
            'non_empty_count': non_empty_count,
        })
    tables.sort(key=lambda x: (x['top_row'], x['left_col']))
    for idx, t in enumerate(tables):
        t['table_id'] = idx
    return tables



def _guess_headers(table_df: 'pd.DataFrame') -> List[str]:
    if table_df is None or table_df.empty:
        return []
    first_row = table_df.iloc[0].tolist()
    headers: List[str] = []
    for i, v in enumerate(first_row):
        txt = _cell_to_display_text(v).strip()
        headers.append(txt if txt else f'col_{i + 1}')
    return headers



def _segment_images_for_table(table_box: Dict[str, Any], sheet_images: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    table_images: List[Dict[str, Any]] = []
    floating_images: List[Dict[str, Any]] = []
    for img in sheet_images or []:
        if img.get('type') == 'cell_image':
            row = img.get('anchor_row')
            col = img.get('anchor_col')
            if row is not None and col is not None:
                if table_box['top_row'] <= row <= table_box['bottom_row'] and table_box['left_col'] <= col <= table_box['right_col']:
                    table_images.append(img)
        else:
            floating_images.append(img)
    return table_images, floating_images



def _build_sheet_unit(sheet_name: str, sheet_index: int, sheet_df: 'pd.DataFrame', sheet_non_empty_df: 'pd.DataFrame',
                      sheet_bbox: Optional[Tuple[int, int, int, int]], formulas: List[Dict[str, Any]],
                      sheet_images: List[Dict[str, Any]]) -> Dict[str, Any]:
    sheet_md = _df_to_markdown(sheet_df)
    unit_text_parts = [f'# Sheet: {sheet_name}']
    if sheet_md:
        unit_text_parts.append(sheet_md)
    else:
        unit_text_parts.append('[EMPTY_SHEET]')
    return {
        'unit_text': '\n\n'.join(unit_text_parts),
        'indent_level': 0,
        'order': sheet_index,
        'structure_chars': [{'type': 'sheet', 'sheet_name': sheet_name}],
        'sheet_name': sheet_name,
        'sheet_index': sheet_index,
        'sheet_df': sheet_df,
        'sheet_values': _to_serializable_df(sheet_df) if sheet_df is not None else [],
        'sheet_markdown': sheet_md,
        'sheet_non_empty_df': sheet_non_empty_df,
        'sheet_non_empty_values': _to_serializable_df(sheet_non_empty_df) if sheet_non_empty_df is not None and not sheet_non_empty_df.empty else [],
        'sheet_non_empty_markdown': _df_to_markdown(sheet_non_empty_df) if sheet_non_empty_df is not None else '',
        'sheet_bbox': sheet_bbox,
        'formula_cells': formulas,
        'sheet_images': sheet_images,
    }



def _sanitize_metadata(metadata: Dict[str, Any]) -> Dict[str, Any]:
    display = {}
    for key, value in metadata.items():
        if key in {'sheets'}:
            summary = []
            for item in value:
                summary.append({
                    'sheet_name': item.get('sheet_name'),
                    'shape': item.get('shape'),
                    'non_empty_shape': item.get('non_empty_shape'),
                    'table_count': item.get('table_count'),
                    'formula_count': item.get('formula_count'),
                    'image_count': item.get('image_count'),
                })
            display[key] = summary
        elif key == 'images' and isinstance(value, list):
            summarized = []
            for img in value:
                summarized.append({
                    'id': img.get('id'),
                    'placeholder': img.get('placeholder'),
                    'sheet_name': img.get('sheet_name'),
                    'anchor_cell': img.get('anchor_cell'),
                    'type': img.get('type'),
                    'size_bytes': img.get('size_bytes'),
                    'has_base64': 'base64' in img,
                    'has_bytes': 'bytes' in img,
                })
            display[key] = summarized
        else:
            display[key] = value
    return display


# ============================================================
# 1. Preview（預覽）
# ============================================================

def preview(
        file_path: str,
        separator: str = '\n\n---\n\n',
        max_rows_per_sheet: Optional[int] = None,
        max_cols_per_sheet: Optional[int] = None,
        **kwargs
    ) -> str:
    _require_dependencies()
    if not os.path.exists(file_path):
        raise FileNotFoundError(f'檔案不存在: {file_path}')

    normalized_path = _normalize_to_xlsx(file_path)
    m_logger.info(f'[preview] 開始預覽檔案: {normalized_path}')

    xls = pd.ExcelFile(normalized_path, engine='openpyxl')
    parts: List[str] = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None, engine='openpyxl')
        part = [f'# Sheet: {sheet_name}']
        part.append(_df_to_markdown(df, max_rows=max_rows_per_sheet, max_cols=max_cols_per_sheet))
        parts.append('\n\n'.join([p for p in part if p]))
    return separator.join(parts)



def parse(file_path: str, **kwargs) -> str:
    return preview(file_path=file_path, **kwargs)


# ============================================================
# 2. Extract（提取）
# ============================================================

def extract(
        file_path: Optional[str] = None,
        text: Optional[str] = None,
        include_images: bool = True,
        image_placeholder: bool = True,
        to_base64: bool = True,
        include_formulas: bool = True,
        **kwargs
    ) -> Dict[str, Any]:
    _require_dependencies()
    if text is not None and not file_path:
        return _extract_from_text(text)
    if not file_path:
        raise ValueError('extract() 需要 file_path 或 text 其中之一')
    if not os.path.exists(file_path):
        raise FileNotFoundError(f'檔案不存在: {file_path}')

    original_file_path = file_path
    normalized_path = _normalize_to_xlsx(file_path)
    m_logger.info(f'[extract] 開始提取 Excel 檔案: {normalized_path}')

    wb = load_workbook(normalized_path, data_only=False)
    metadata: Dict[str, Any] = {
        'file_path': original_file_path,
        'processed_file_path': normalized_path,
        'source_format': Path(original_file_path).suffix.lower().lstrip('.'),
        'sheet_count': len(wb.worksheets),
        'sheet_names': [ws.title for ws in wb.worksheets],
        'sheets': [],
        'images': [],
        'formula_cells': [],
        'has_images': False,
        'has_formulas': False,
        'image_count': 0,
        'formula_count': 0,
        'total_tables': 0,
    }
    unit_paras: List[Dict[str, Any]] = []

    for sheet_index, ws in enumerate(wb.worksheets):
        sheet_df = _worksheet_to_dataframe(ws, include_formulas=include_formulas)
        sheet_non_empty_df, sheet_bbox = _trim_empty_df(sheet_df)
        formulas = _extract_formula_cells(ws) if include_formulas else []
        sheet_images = _extract_sheet_images(ws, include_images=include_images and image_placeholder, to_base64=to_base64)
        if not image_placeholder:
            for img in sheet_images:
                img.pop('placeholder', None)
                img.pop('id', None)

        sheet_unit = _build_sheet_unit(
            sheet_name=ws.title,
            sheet_index=sheet_index,
            sheet_df=sheet_df,
            sheet_non_empty_df=sheet_non_empty_df,
            sheet_bbox=sheet_bbox,
            formulas=formulas,
            sheet_images=sheet_images,
        )
        unit_paras.append(sheet_unit)

        tables = _detect_tables_from_sheet_df(sheet_non_empty_df)
        # 將未落在任何表格內的圖像獨立成 unit_para（含浮動圖、表格外的 cell_image）
        if sheet_images:
            floating_images: List[Dict[str, Any]] = []
            for img in sheet_images:
                placeholder = img.get('placeholder')
                if not placeholder:
                    continue
                img_type = img.get('type')
                in_table = False
                if img_type == 'cell_image':
                    row = img.get('anchor_row')
                    col = img.get('anchor_col')
                    if row is not None and col is not None:
                        for t in tables:
                            if t['top_row'] <= row <= t['bottom_row'] and t['left_col'] <= col <= t['right_col']:
                                in_table = True
                                break
                if not in_table:
                    floating_images.append(img)

            for img in floating_images:
                unit_paras.append({
                    'unit_text': '\n'.join([f'# Sheet: {ws.title}', img.get('placeholder', '')]),
                    'indent_level': 0,
                    'order': len(unit_paras),
                    'unit_type': 'sheet_image',
                    'sheet_name': ws.title,
                    'sheet_index': sheet_index,
                    'images': [img],
                    'sheet_images': [img],
                    'structure_chars': [
                        {'type': 'sheet', 'sheet_name': ws.title},
                        {'type': 'image', 'image_id': img.get('id'), 'anchor_cell': img.get('anchor_cell')},
                    ],
                })
        metadata['total_tables'] += len(tables)
        metadata['formula_cells'].extend(formulas)
        metadata['images'].extend(sheet_images)
        metadata['sheets'].append({
            'sheet_name': ws.title,
            'sheet_index': sheet_index,
            'shape': [int(sheet_df.shape[0]), int(sheet_df.shape[1])],
            'non_empty_shape': [int(sheet_non_empty_df.shape[0]), int(sheet_non_empty_df.shape[1])] if sheet_non_empty_df is not None and not sheet_non_empty_df.empty else [0, 0],
            'sheet_bbox': sheet_bbox,
            'table_count': len(tables),
            'formula_count': len(formulas),
            'image_count': len(sheet_images),
            'tables_preview': [{
                'table_id': t['table_id'],
                'top_row': t['top_row'],
                'left_col': t['left_col'],
                'bottom_row': t['bottom_row'],
                'right_col': t['right_col'],
                'shape': t['shape'],
            } for t in tables],
        })

    metadata['formula_count'] = len(metadata['formula_cells'])
    metadata['image_count'] = len(metadata['images'])
    metadata['has_images'] = metadata['image_count'] > 0
    metadata['has_formulas'] = metadata['formula_count'] > 0

    m_logger.info(f"[extract] 提取完成: sheet數={metadata['sheet_count']}, table數={metadata['total_tables']}, 公式數={metadata['formula_count']}, 圖像數={metadata['image_count']}")
    return {
        'unit_paras': unit_paras,
        'metadata': metadata,
    }



def _extract_from_text(text: str, **kwargs) -> Dict[str, Any]:
    df = pd.DataFrame([line.split('\t') for line in text.splitlines() if line.strip()])
    unit = {
        'unit_text': _df_to_markdown(df),
        'indent_level': 0,
        'order': 0,
        'structure_chars': [{'type': 'sheet', 'sheet_name': 'TextSheet'}],
        'sheet_name': 'TextSheet',
        'sheet_index': 0,
        'sheet_df': df,
        'sheet_values': _to_serializable_df(df),
        'sheet_markdown': _df_to_markdown(df),
        'sheet_non_empty_df': df,
        'sheet_non_empty_values': _to_serializable_df(df),
        'sheet_non_empty_markdown': _df_to_markdown(df),
        'sheet_bbox': (1, 1, int(df.shape[0]), int(df.shape[1]) if len(df.shape) > 1 else 1),
        'formula_cells': [],
        'sheet_images': [],
    }
    return {
        'unit_paras': [unit],
        'metadata': {
            'file_path': None,
            'processed_file_path': None,
            'source_format': 'text',
            'sheet_count': 1,
            'sheet_names': ['TextSheet'],
            'sheets': [],
            'images': [],
            'formula_cells': [],
            'has_images': False,
            'has_formulas': False,
            'image_count': 0,
            'formula_count': 0,
            'total_tables': 1,
        }
    }


# ============================================================
# 3. Segment（分段）
# ============================================================

def segment(
        unit_paras: List[Dict[str, Any]],
        suitable_char_count: int = 500,
        separator: str = '\n',
        metadata: Dict[str, Any] = None,
        enable_image_llm: bool = True,
        llm_provider: str = 'remote',
        llm_model: str = 'remote8b',
        llm_base_url: str = None,
        image_context_window: int = 200,
        max_images_per_batch: int = 50,
        image_prompt_template: str = None,
        image_llm_provider: str = 'openai',
        image_llm_model: str = 'gpt4o_chat',
        min_non_empty_cells: int = 1,
        **kwargs
    ) -> List[Dict[str, Any]]:
    if not unit_paras:
        m_logger.warning('[segment] unit_paras 為空，返回空列表')
        return []

    results: List[Dict[str, Any]] = []
    order = 0
    for unit in unit_paras:
        sheet_name = unit.get('sheet_name', '')
        if unit.get('unit_type') == 'sheet_image':
            results.append({
                'segment_id': f'seg_{order:04d}',
                'unit_text': unit.get('unit_text', ''),
                'text': unit.get('unit_text', ''),
                'indent_level': 0,
                'order': order,
                'sheet_name': sheet_name,
                'sheet_index': unit.get('sheet_index', 0),
                'segment_type': 'sheet_image',
                'top_row': None,
                'left_col': None,
                'bottom_row': None,
                'right_col': None,
                'table_df': pd.DataFrame(),
                'table_values': [],
                'table_markdown': '',
                'headers': [],
                'images': unit.get('images', []),
                'sheet_images': unit.get('sheet_images', []),
                'structure_chars': unit.get('structure_chars', [{'type': 'sheet', 'sheet_name': sheet_name}]),
            })
            order += 1
            continue
        sheet_df = unit.get('sheet_non_empty_df') if unit.get('sheet_non_empty_df') is not None else unit.get('sheet_df')
        if sheet_df is None or getattr(sheet_df, 'empty', True):
            results.append({
                'segment_id': f'seg_{order:04d}',
                'unit_text': unit.get('unit_text', ''),
                'text': unit.get('unit_text', ''),
                'indent_level': 0,
                'order': order,
                'sheet_name': sheet_name,
                'sheet_index': unit.get('sheet_index', 0),
                'segment_type': 'empty_sheet',
                'top_row': None,
                'left_col': None,
                'bottom_row': None,
                'right_col': None,
                'table_df': pd.DataFrame(),
                'table_values': [],
                'table_markdown': '',
                'headers': [],
                'images': unit.get('sheet_images', []),
                'sheet_images': unit.get('sheet_images', []),
                'structure_chars': [{'type': 'sheet', 'sheet_name': sheet_name}],
            })
            order += 1
            continue

        tables = _detect_tables_from_sheet_df(sheet_df, min_non_empty_cells=min_non_empty_cells)
        if not tables:
            tables = [{
                'table_id': 0,
                'top_row': 1,
                'left_col': 1,
                'bottom_row': int(sheet_df.shape[0]),
                'right_col': int(sheet_df.shape[1]),
                'shape': [int(sheet_df.shape[0]), int(sheet_df.shape[1])],
                'table_df': sheet_df.copy(),
                'table_values': _to_serializable_df(sheet_df),
                'table_markdown': _df_to_markdown(sheet_df),
                'non_empty_count': int(sheet_df.applymap(lambda x: _normalize_none_like(x) is not None).to_numpy().sum()),
            }]

        for table in tables:
            table_images, sheet_images = _segment_images_for_table(table, unit.get('sheet_images', []))
            headers = _guess_headers(table['table_df'])
            structure_chars = [
                {'type': 'sheet', 'sheet_name': sheet_name},
                {
                    'type': 'table',
                    'table_id': table['table_id'],
                    'top_row': table['top_row'],
                    'left_col': table['left_col'],
                    'bottom_row': table['bottom_row'],
                    'right_col': table['right_col'],
                }
            ]
            for img in table_images:
                if img.get('placeholder'):
                    structure_chars.append({'type': 'image', 'image_id': img.get('id'), 'anchor_cell': img.get('anchor_cell')})
            text_parts = [f'# Sheet: {sheet_name}', f'## Table {table["table_id"]}', table['table_markdown']]
            if table_images:
                image_line = ' '.join([img.get('placeholder', f'[IMAGE_{img.get("local_index", 0)}]') for img in table_images])
                text_parts.append(image_line)
            seg_text = separator.join([p for p in text_parts if p])
            results.append({
                'segment_id': f'seg_{order:04d}',
                'unit_text': seg_text,
                'text': seg_text,
                'indent_level': 0,
                'order': order,
                'sheet_name': sheet_name,
                'sheet_index': unit.get('sheet_index', 0),
                'segment_type': 'table',
                'table_id': table['table_id'],
                'top_row': table['top_row'],
                'left_col': table['left_col'],
                'bottom_row': table['bottom_row'],
                'right_col': table['right_col'],
                'table_df': table['table_df'],
                'table_values': table['table_values'],
                'table_markdown': table['table_markdown'],
                'headers': headers,
                'images': table_images,
                'sheet_images': sheet_images,
                'formula_cells': [f for f in unit.get('formula_cells', []) if _formula_in_table(f, table)],
                'structure_chars': structure_chars,
            })
            order += 1

    m_logger.info(f'[segment] 分段完成: {len(results)} 個 table segments')

    if enable_image_llm and metadata and metadata.get('has_images'):
        _analyze_images_via_batch(
            segments=results,
            metadata=metadata,
            llm_provider=image_llm_provider,
            llm_model=image_llm_model,
            llm_base_url=llm_base_url,
            enable_image_llm=enable_image_llm,
            image_context_window=image_context_window,
            max_images_per_batch=max_images_per_batch,
            image_prompt_template=image_prompt_template,
        )
    return results



def _formula_in_table(formula_info: Dict[str, Any], table: Dict[str, Any]) -> bool:
    cell = formula_info.get('cell')
    if not cell:
        return False
    m = re.match(r'([A-Z]+)(\d+)$', str(cell))
    if not m:
        return False
    letters, row_txt = m.groups()
    row = int(row_txt)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - 64)
    return table['top_row'] <= row <= table['bottom_row'] and table['left_col'] <= col <= table['right_col']



def _analyze_images_via_batch(
        segments: List[Dict[str, Any]],
        metadata: Dict[str, Any],
        llm_provider: str = 'openai',
        llm_model: str = 'gpt4o_chat',
        llm_base_url: str = None,
        enable_image_llm: bool = True,
        image_context_window: int = 200,
        max_images_per_batch: int = 50,
        image_prompt_template: str = None,
        **kwargs
    ) -> Dict[str, Any]:
    if analyze_images_via_batch_common is None:
        m_logger.warning('[_analyze_images_via_batch] shared helper unavailable')
        return {}
    return analyze_images_via_batch_common(
        segments=segments,
        metadata=metadata,
        llm_provider=llm_provider,
        llm_model=llm_model,
        llm_base_url=llm_base_url,
        enable_image_llm=enable_image_llm,
        image_context_window=image_context_window,
        max_images_per_batch=max_images_per_batch,
        image_prompt_template=image_prompt_template,
        placeholder_pattern=r'\[IMAGE_PLACEHOLDER_(\d+)\]',
        placeholder_replacements=['[IMAGE_PLACEHOLDER_{image_id}]'],
        text_keys=['unit_text', 'call_prompt', 'text'],
        config=m_config,
        logger=m_logger
    )


# ============================================================
# 4. Chunk（分塊）
# ============================================================


class BaseChunker:
    def __init__(self, parse_mode: str, max_cell_chars: int, digit: int, metadata: Optional[Dict[str, Any]], skip_dedup: bool = False, skip_keyword_gen: bool = False):
        self.parse_mode = parse_mode
        self.max_cell_chars = max_cell_chars
        self.digit = digit
        self.metadata = metadata
        self.skip_dedup = skip_dedup
        self.skip_keyword_gen = skip_keyword_gen
        self._headers_cache: Dict[int, List[str]] = {}  # 快取 headers（使用 DataFrame id 作為鍵）

    def _format_cell(self, value: Any) -> str:
        return _cell_to_display_text(value, max_len=self.max_cell_chars, digit=self.digit)

    def _build(self, chunk_text: str, seg: Dict[str, Any], chunk_order: int,
               row_index: Optional[int], col_index: Optional[int], cell_ref: Optional[str]) -> Dict[str, Any]:
        return _build_chunk_from_text(
            chunk_text=chunk_text,
            seg=seg,
            parse_mode=self.parse_mode,
            chunk_order=chunk_order,
            metadata=self.metadata,
            row_index=row_index,
            col_index=col_index,
            cell_ref=cell_ref,
            skip_dedup=self.skip_dedup,
            skip_keyword_gen=self.skip_keyword_gen,
        )

    def _get_headers(self, table_df: 'pd.DataFrame', cached_headers: Optional[List[str]] = None) -> List[str]:
        """取得 headers，優先使用快取的"""
        if cached_headers is not None:
            return cached_headers
        # 使用 DataFrame 的 id 作為快取鍵
        df_id = id(table_df)
        if df_id in self._headers_cache:
            return self._headers_cache[df_id]
        headers = _guess_headers(table_df)
        self._headers_cache[df_id] = headers
        return headers

    def generate(
        self,
        seg: Dict[str, Any],
        table_df: 'pd.DataFrame',
        chunk_order: int,
        index_signals: Optional[Dict[str, Any]] = None,
        cached_headers: Optional[List[str]] = None
    ) -> Tuple[List[Dict[str, Any]], int]:
        raise NotImplementedError


class RowChunker(BaseChunker):
    def generate(
        self,
        seg: Dict[str, Any],
        table_df: 'pd.DataFrame',
        chunk_order: int,
        index_signals: Optional[Dict[str, Any]] = None,
        cached_headers: Optional[List[str]] = None
    ) -> Tuple[List[Dict[str, Any]], int]:
        signals = index_signals or {}
        index_row = bool(signals.get('index_row', False))
        index_col = bool(signals.get('index_col', False))
        headers = signals.get('row_headers') if index_row else self._get_headers(table_df, cached_headers)
        if index_row and not headers:
            headers = [
                _cell_to_display_text(v, max_len=self.max_cell_chars, digit=self.digit).strip()
                for v in table_df.iloc[0].tolist()
            ]
        headers = headers or []
        chunks: List[Dict[str, Any]] = []
        top_row = seg.get('top_row', 1)
        row_start = 1 if index_row else 0
        col_start = 1 if index_col else 0
        for ridx in range(row_start, table_df.shape[0]):
            values = table_df.iloc[ridx].tolist()
            parts = []
            for i in range(col_start, len(values)):
                header_name = headers[i] if i < len(headers) else ''
                if header_name is not None:
                    header_name = str(header_name)
                if index_row and header_name.startswith('col_'):
                    header_name = ''
                if not header_name:
                    continue
                cell_text = self._format_cell(values[i])
                if cell_text != '':
                    parts.append(f'{header_name}={cell_text}')
            if not parts:
                continue
            row_text = ' | '.join(parts)
            actual_row = top_row + ridx
            chunks.append(self._build(row_text, seg, chunk_order, actual_row, None, None))
            chunk_order += 1
        return chunks, chunk_order


class ColumnChunker(BaseChunker):
    def generate(
        self,
        seg: Dict[str, Any],
        table_df: 'pd.DataFrame',
        chunk_order: int,
        index_signals: Optional[Dict[str, Any]] = None,
        cached_headers: Optional[List[str]] = None
    ) -> Tuple[List[Dict[str, Any]], int]:
        signals = index_signals or {}
        index_row = bool(signals.get('index_row', False))
        index_col = bool(signals.get('index_col', False))
        headers = signals.get('row_headers') if index_row else self._get_headers(table_df, cached_headers)
        if index_row and not headers:
            headers = [
                _cell_to_display_text(v, max_len=self.max_cell_chars, digit=self.digit).strip()
                for v in table_df.iloc[0].tolist()
            ]
        headers = headers or []
        chunks: List[Dict[str, Any]] = []
        left_col = seg.get('left_col', 1)

        if index_col:
            row_start = 1 if index_row else 0
            row_headers = (signals.get('col_headers') or [])[row_start:]
            if not row_headers:
                col_values = table_df.iloc[row_start:, 0].tolist()
                row_headers = [
                    _cell_to_display_text(v, max_len=self.max_cell_chars, digit=self.digit).strip()
                    for v in col_values
                ]
            for cidx in range(1, table_df.shape[1]):
                values = table_df.iloc[row_start:, cidx].tolist()
                parts = []
                for r_i, v in enumerate(values):
                    header_name = row_headers[r_i] if r_i < len(row_headers) else ''
                    if header_name is not None:
                        header_name = str(header_name)
                    if not header_name:
                        continue
                    cell_text = self._format_cell(v)
                    if cell_text != '':
                        parts.append(f'{header_name}={cell_text}')
                if not parts:
                    continue
                col_name = headers[cidx] if cidx < len(headers) and headers[cidx] else f'col_{cidx + 1}'
                col_text = f'{col_name}: ' + ' | '.join(parts)
                actual_col = left_col + cidx
                chunks.append(self._build(col_text, seg, chunk_order, None, actual_col, None))
                chunk_order += 1
            return chunks, chunk_order

        for cidx in range(table_df.shape[1]):
            values = table_df.iloc[:, cidx].tolist()
            non_empty_values = []
            for v in values:
                cell_text = self._format_cell(v)
                if cell_text != '':
                    non_empty_values.append(cell_text)
            if not non_empty_values:
                continue
            col_name = headers[cidx] if cidx < len(headers) else f'col_{cidx + 1}'
            col_text = f'{col_name}: ' + ' | '.join(non_empty_values)
            actual_col = left_col + cidx
            chunks.append(self._build(col_text, seg, chunk_order, None, actual_col, None))
            chunk_order += 1
        return chunks, chunk_order


class CellChunker(BaseChunker):
    def generate(
        self,
        seg: Dict[str, Any],
        table_df: 'pd.DataFrame',
        chunk_order: int,
        index_signals: Optional[Dict[str, Any]] = None,
        cached_headers: Optional[List[str]] = None
    ) -> Tuple[List[Dict[str, Any]], int]:
        headers = self._get_headers(table_df, cached_headers)
        chunks: List[Dict[str, Any]] = []
        top_row = seg.get('top_row', 1)
        left_col = seg.get('left_col', 1)
        for ridx in range(table_df.shape[0]):
            for cidx in range(table_df.shape[1]):
                val = table_df.iat[ridx, cidx]
                text = self._format_cell(val)
                if text == '':
                    continue
                actual_row = top_row + ridx
                actual_col = left_col + cidx
                cell_ref = f'{_colnum_to_letter(actual_col)}{actual_row}'
                header = headers[cidx] if cidx < len(headers) else f'col_{actual_col}'
                cell_text = f'{cell_ref} | {header}={text}'
                chunks.append(self._build(cell_text, seg, chunk_order, actual_row, actual_col, cell_ref))
                chunk_order += 1
        return chunks, chunk_order


class PreviewChunker(BaseChunker):
    def generate(
        self,
        seg: Dict[str, Any],
        table_df: 'pd.DataFrame',
        chunk_order: int,
        index_signals: Optional[Dict[str, Any]] = None,
        cached_headers: Optional[List[str]] = None
    ) -> Tuple[List[Dict[str, Any]], int]:
        if table_df is None or getattr(table_df, 'empty', True):
            return [], chunk_order
        try:
            view_df = table_df.copy()
            view_df = view_df.applymap(lambda v: self._format_cell(v))
            table_md = _df_to_markdown(view_df)
        except Exception:
            table_md = _df_to_markdown(table_df)
        chunks = [self._build(table_md, seg, chunk_order, None, None, None)]
        return chunks, chunk_order + 1


class DiagonalChunker(BaseChunker):
    def generate(
        self,
        seg: Dict[str, Any],
        table_df: 'pd.DataFrame',
        chunk_order: int,
        index_signals: Optional[Dict[str, Any]] = None,
        cached_headers: Optional[List[str]] = None
    ) -> Tuple[List[Dict[str, Any]], int]:
        if table_df is None or getattr(table_df, 'empty', True):
            return [], chunk_order
        chunks: List[Dict[str, Any]] = []
        rows, cols = table_df.shape
        top_row = seg.get('top_row', 1)
        left_col = seg.get('left_col', 1)
        for diag in range(rows + cols - 1):
            parts = []
            for ridx in range(rows):
                cidx = diag - ridx
                if cidx < 0 or cidx >= cols:
                    continue
                val = table_df.iat[ridx, cidx]
                text = self._format_cell(val)
                if text == '':
                    continue
                actual_row = top_row + ridx
                actual_col = left_col + cidx
                cell_ref = f'{_colnum_to_letter(actual_col)}{actual_row}'
                parts.append(f'{cell_ref}={text}')
            if not parts:
                continue
            diag_text = f'diag_{diag + 1}: ' + ' | '.join(parts)
            chunks.append(self._build(diag_text, seg, chunk_order, None, None, None))
            chunk_order += 1
        return chunks, chunk_order


def _parse_mode_and_reason_from_llm_text(text: str) -> Tuple[Optional[str], Optional[str]]:
    if not text:
        return None, None
    stripped = text.strip()
    # 嘗試解析 JSON 回應
    try:
        data = json.loads(stripped)
        if isinstance(data, dict):
            mode = (data.get('mode') or data.get('parse_mode') or '').strip().lower()
            reason = data.get('reason')
            if mode in {'row', 'column', 'preview', 'diagonal', 'cell'}:
                return mode, reason
    except Exception:
        pass
    lowered = stripped.lower()
    mode = None
    for cand in ('row', 'column', 'preview', 'diagonal', 'cell'):
        if cand in lowered:
            mode = cand
            break
    reason = None
    # 嘗試從「理由:」或「原因:」之後擷取
    for key in ('理由', '原因', 'reason'):
        if key in lowered:
            idx = lowered.find(key)
            reason = stripped[idx + len(key):].lstrip('：: \n')
            break
    return mode, reason


def _clean_llm_json_object(text: str) -> str:
    """
        清理 LLM 輸出，擷取第一個完整 JSON 物件字串。
        舉例: '{"a":1}<|eot_id|>...' -> '{"a":1}'
    """
    if not text or not isinstance(text, str):
        return ''
    cleaned = text.strip()
    special_tokens = ['<|eot_id|>', '<|end_of_text|>', '<|endoftext|>', '<|im_end|>', '<|end|>']
    for token in special_tokens:
        cleaned = cleaned.replace(token, '')
    cleaned = cleaned.strip()

    # Prefer JSON inside code block if present
    code_block_match = re.search(r'```(?:json)?\s*(\{[\s\S]*?\})\s*```', cleaned, re.DOTALL)
    if code_block_match:
        return code_block_match.group(1).strip()

    start = cleaned.find('{')
    if start == -1:
        return ''

    brace_count = 0
    end = -1
    for i in range(start, len(cleaned)):
        ch = cleaned[i]
        if ch == '{':
            brace_count += 1
        elif ch == '}':
            brace_count -= 1
            if brace_count == 0:
                end = i
                break
    if end != -1:
        return cleaned[start:end + 1].strip()

    # Fallback: try to use last closing brace
    end = cleaned.rfind('}')
    if end != -1 and end > start:
        return cleaned[start:end + 1].strip()
    return ''


def _normalize_header_list(headers: Any, max_items: int = 12) -> List[Any]:
    if not isinstance(headers, list):
        return []
    cleaned = [h for h in headers if h not in (None, '')]
    if max_items and len(cleaned) > max_items:
        cleaned = cleaned[:max_items]
    return cleaned


def _is_sequential_numeric_list(headers: List[Any]) -> bool:
    nums: List[int] = []
    for v in headers:
        if isinstance(v, bool):
            return False
        if isinstance(v, int):
            nums.append(int(v))
            continue
        if isinstance(v, str) and re.fullmatch(r'-?\d+', v.strip()):
            nums.append(int(v.strip()))
            continue
        return False
    if len(nums) < 3:
        return False
    nums_sorted = sorted(nums)
    if len(nums_sorted) != len(set(nums_sorted)):
        return False
    return nums_sorted[-1] - nums_sorted[0] == len(nums_sorted) - 1


def _extract_bool_field(text: str, key: str) -> Optional[bool]:
    m = re.search(rf'"{re.escape(key)}"\s*:\s*(true|false)', text, re.IGNORECASE)
    if not m:
        return None
    return m.group(1).lower() == 'true'


def _extract_array_field(text: str, key: str) -> Optional[List[Any]]:
    key_pos = text.find(f'"{key}"')
    if key_pos == -1:
        return None
    start = text.find('[', key_pos)
    if start == -1:
        return None
    in_str = False
    escape = False
    depth = 0
    end = -1
    for i in range(start, len(text)):
        ch = text[i]
        if in_str:
            if escape:
                escape = False
            elif ch == '\\':
                escape = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
            continue
        if ch == '[':
            depth += 1
            continue
        if ch == ']':
            depth -= 1
            if depth == 0:
                end = i
                break
    if end == -1:
        return None
    arr_text = text[start:end + 1]
    for marker in ('"reason":', '"mode":', '"index_row":', '"index_col":'):
        if marker in arr_text:
            prefix = arr_text.split(marker)[0]
            prefix = re.sub(r',\s*$', '', prefix)
            arr_text = prefix + ']'
            break
    try:
        val = json.loads(arr_text)
        return val if isinstance(val, list) else None
    except Exception:
        items = []
        inner = arr_text.strip().lstrip('[').rstrip(']')
        for part in inner.split(','):
            part = part.strip()
            if not part:
                continue
            if part.startswith('"') and part.endswith('"'):
                items.append(part[1:-1])
            elif re.fullmatch(r'-?\d+', part):
                items.append(int(part))
            else:
                items.append(part)
        return items


def _parse_index_relaxed(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None
    index_row = _extract_bool_field(text, 'index_row')
    index_col = _extract_bool_field(text, 'index_col')
    row_headers = _extract_array_field(text, 'row_headers') or []
    col_headers = _extract_array_field(text, 'col_headers') or []
    reason_match = re.search(r'"reason"\s*:\s*"(.*?)"', text, re.DOTALL)
    reason = reason_match.group(1) if reason_match else None
    if index_row is None and index_col is None and not row_headers and not col_headers:
        return None
    index_row = bool(index_row)
    index_col = bool(index_col)
    row_headers = _normalize_header_list(row_headers)
    col_headers = _normalize_header_list(col_headers)
    if _is_sequential_numeric_list(row_headers):
        row_headers = []
        index_row = False
    if _is_sequential_numeric_list(col_headers):
        col_headers = []
        index_col = False
    if index_row and not row_headers:
        index_row = False
    if index_col and not col_headers:
        index_col = False
    return {
        'index_row': index_row,
        'index_col': index_col,
        'row_headers': row_headers,
        'col_headers': col_headers,
        'reason': reason,
    }


def _build_llm_table_prompt(
    table_md: str,
    header_signals: Optional[Dict[str, Any]] = None,
    numeric_ratio: Optional[float] = None
) -> str:
    signals_text = ''
    if header_signals:
        row_headers = header_signals.get('row_headers') or []
        col_headers = header_signals.get('col_headers') or []
        if len(row_headers) > 8:
            row_headers = row_headers[:8] + ['...']
        if len(col_headers) > 8:
            col_headers = col_headers[:8] + ['...']
        signals_text = (
            "自動偵測：\n"
            f"- index_row={header_signals.get('index_row')}, row_header_score={header_signals.get('row_header_score')}\n"
            f"- index_col={header_signals.get('index_col')}, col_header_score={header_signals.get('col_header_score')}\n"
            f"- row_headers={row_headers}\n"
            f"- col_headers={col_headers}\n"
        )
    if numeric_ratio is not None:
        signals_text += f"- numeric_ratio={numeric_ratio:.3f} (numeric_cells / non_empty_cells)\n"
    return (
        "你是資料整理的助手，請根據下列表格內容選擇最適合的記錄模式。\n"
        "可選模式：row、column、preview、diagonal。\n"
        "判斷依據：\n"
        "- row：在表格前幾 row 有明顯的標題式 header，對應下來每一 col 與 header 的值域屬性相關。\n"
        "- column：在表格前幾 column 有明顯的標題式 header，對應下來每一 column 與 header 的值域屬性相關。\n"
        "- diagonal：前幾 row、前幾 column 找的到完整的標題式名稱，對應每一個儲存格都是相關值域；或者沒有 row/column 標題式名稱。\n"
        "- preview：前幾種都不適合；若 numeric_ratio > 0.5，優先選 preview。\n"
        "請用 JSON 回覆：{\"mode\":\"<row|column|preview|diagonal>\",\"reason\":\"...\"}。\n\n"
        f"{signals_text}\n"
        f"表格內容：\n{table_md}\n"
    )


def _build_llm_index_prompt(table_md: str) -> str:
    return (
        "你是資料整理的助手，請判斷下列表格是否有明確的 index_row 與 index_col。\n"
        "定義：\n"
        "- index_row：前幾 row 有明顯的標題式 header，對應到每一欄的值域屬性。\n"
        "- index_col：前幾 column 有明顯的標題式 header，對應到每一列的值域屬性。\n"
        "注意：\n"
        "- 若第一列/欄只是連續數字（例如 0,1,2... 或 1,2,3...）、空白或純序號，視為 DataFrame 自動 index/欄位序號，不要當成 header。\n"
        "- row_headers / col_headers 應回傳有語意的標籤，不要回傳一長串連續整數。\n"
        "- row_headers / col_headers 最多列出 12 個即可。\n"
        "請用 JSON 回覆："
        "{\"index_row\":true/false,\"index_col\":true/false,"
        "\"row_headers\":[...],\"col_headers\":[...],\"reason\":\"...\"}\n\n"
        f"表格內容：\n{table_md}\n"
    )


def _build_llm_combined_prompt(
        table_md: str,
        header_signals: Optional[Dict[str, Any]] = None,
        numeric_ratio: Optional[float] = None
    ) -> str:
    """合併的 prompt：同時要求 index 和 mode 偵測"""
    signals_text = ''
    if header_signals:
        row_headers = header_signals.get('row_headers') or []
        col_headers = header_signals.get('col_headers') or []
        if len(row_headers) > 8:
            row_headers = row_headers[:8] + ['...']
        if len(col_headers) > 8:
            col_headers = col_headers[:8] + ['...']
        signals_text = (
            "自動偵測：\n"
            f"- index_row={header_signals.get('index_row')}, row_header_score={header_signals.get('row_header_score')}\n"
            f"- index_col={header_signals.get('index_col')}, col_header_score={header_signals.get('col_header_score')}\n"
            f"- row_headers={row_headers}\n"
            f"- col_headers={col_headers}\n"
        )
    if numeric_ratio is not None:
        signals_text += f"- numeric_ratio={numeric_ratio:.3f} (numeric_cells / non_empty_cells)\n"
        signals_text += "- if numeric_ratio > 0.5, prefer preview\n"
    return (
        "你是資料整理的助手，請同時判斷下列表格的 index 與最適合的記錄模式。\n\n"
        "【Index 判斷】\n"
        "- index_row：前幾 row 有明顯的標題式 header，對應到每一欄的值域屬性。\n"
        "- index_col：前幾 column 有明顯的標題式 header，對應到每一列的值域屬性。\n\n"
        "注意：\n"
        "- 若第一列/欄只是連續數字（例如 0,1,2... 或 1,2,3...）、空白或純序號，視為 DataFrame 自動 index/欄位序號，不要當成 header。\n"
        "- row_headers / col_headers 應回傳有語意的標籤，不要回傳一長串連續整數。\n\n"
        "- row_headers / col_headers 最多列出 12 個即可。\n\n"
        "【Mode 判斷】\n"
        "可選模式：row、column、cell、preview、diagonal。\n"
        "判斷依據：\n"
        "- row：在表格前幾 row 有明顯的標題式 header，且每一欄與 header 的值域屬性對應。\n"
        "- column：在表格前幾 column 有明顯的標題式 header，且每一欄與 header 的值域屬性對應。\n"
        "- cell：需要逐個儲存格獨立記錄，或表格結構不規則無法用 row/column 模式處理。\n"
        "- diagonal：表格資料按對角線方向組織（從左上到右下），同一對角線上的儲存格屬於同一類別或相關聯。\n"
        "- preview：以上皆不適合，直接將整個表格轉為 Markdown 格式。\n\n"
        f"{signals_text}\n"
        "請用 JSON 回覆："
        "{\"index_row\":true/false,\"index_col\":true/false,"
        "\"row_headers\":[...],\"col_headers\":[...],"
        "\"mode\":\"<row|column|cell|preview|diagonal>\",\"reason\":\"...\"}\n\n"
        f"表格內容：\n{table_md}\n"
    )


def _parse_index_from_llm_text(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None
    stripped = _clean_llm_json_object(text)
    if not stripped:
        return None
    try:
        data = json.loads(stripped)
        if not isinstance(data, dict):
            return None
        index_row = bool(data.get('index_row', False))
        index_col = bool(data.get('index_col', False))
        row_headers = data.get('row_headers') or []
        col_headers = data.get('col_headers') or []
        if not isinstance(row_headers, list):
            row_headers = []
        if not isinstance(col_headers, list):
            col_headers = []
        row_headers = _normalize_header_list(row_headers)
        col_headers = _normalize_header_list(col_headers)
        if _is_sequential_numeric_list(row_headers):
            row_headers = []
            index_row = False
        if _is_sequential_numeric_list(col_headers):
            col_headers = []
            index_col = False
        if index_row and not row_headers:
            index_row = False
        if index_col and not col_headers:
            index_col = False
        reason = data.get('reason')
        return {
            'index_row': index_row,
            'index_col': index_col,
            'row_headers': row_headers,
            'col_headers': col_headers,
            'reason': reason,
        }
    except Exception:
        return _parse_index_relaxed(stripped) or _parse_index_relaxed(text)


def _parse_combined_from_llm_text(text: str) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """解析合併的 LLM 回應，返回 (index_dict, mode)"""
    if not text:
        return None, None
    stripped = _clean_llm_json_object(text)
    if not stripped:
        return None, None
    try:
        data = json.loads(stripped)
        if not isinstance(data, dict):
            return None, None
        
        # 解析 index
        index_row = bool(data.get('index_row', False))
        index_col = bool(data.get('index_col', False))
        row_headers = data.get('row_headers') or []
        col_headers = data.get('col_headers') or []
        if not isinstance(row_headers, list):
            row_headers = []
        if not isinstance(col_headers, list):
            col_headers = []
        row_headers = _normalize_header_list(row_headers)
        col_headers = _normalize_header_list(col_headers)
        if _is_sequential_numeric_list(row_headers):
            row_headers = []
            index_row = False
        if _is_sequential_numeric_list(col_headers):
            col_headers = []
            index_col = False
        if index_row and not row_headers:
            index_row = False
        if index_col and not col_headers:
            index_col = False
        reason = data.get('reason')
        index_dict = {
            'index_row': index_row,
            'index_col': index_col,
            'row_headers': row_headers,
            'col_headers': col_headers,
            'reason': reason,
        }
        
        # 解析 mode
        mode = (data.get('mode') or '').strip().lower()
        if mode not in {'row', 'column', 'preview', 'diagonal', 'cell'}:
            mode = None
        
        return index_dict, mode
    except Exception:
        index_dict = _parse_index_relaxed(stripped) or _parse_index_relaxed(text)
        mode = None
        lowered = stripped.lower() if stripped else (text.lower() if isinstance(text, str) else '')
        for cand in ('row', 'column', 'preview', 'diagonal', 'cell'):
            if cand in lowered:
                mode = cand
                break
        return index_dict, mode


def _is_numeric_text(text: str) -> bool:
    if not text:
        return False
    candidate = text.strip().replace(',', '')
    if candidate == '':
        return False
    try:
        float(candidate)
        return True
    except Exception:
        return False


def _is_text_like(value: Any, text: str) -> bool:
    if not text:
        return False
    if isinstance(value, str):
        return any(ch.isalpha() or ('\u4e00' <= ch <= '\u9fff') for ch in text)
    return not _is_numeric_text(text)


def _compute_numeric_cell_ratio(table_df: 'pd.DataFrame', max_cell_chars: int, digit: int) -> float:
    if table_df is None or getattr(table_df, 'empty', True):
        return 0.0
    total = 0
    numeric = 0
    for row in table_df.values.tolist():
        for value in row:
            cell_text = _cell_to_display_text(value, max_len=max_cell_chars, digit=digit).strip()
            if cell_text == '':
                continue
            total += 1
            if _is_numeric_text(cell_text):
                numeric += 1
    if total == 0:
        return 0.0
    return numeric / total


def _analyze_header_signals(
        table_df: 'pd.DataFrame',
        max_cell_chars: int,
        digit: int,
        min_non_empty_cells: int = 1,
        header_threshold: float = 0.3,
    ) -> Dict[str, Any]:
    if table_df is None or getattr(table_df, 'empty', True):
        return {
            'index_row': False,
            'index_col': False,
            'row_header_score': 0.0,
            'col_header_score': 0.0,
            'row_headers': [],
            'col_headers': [],
        }

    rows, cols = table_df.shape
    if rows == 0 or cols == 0:
        return {
            'index_row': False,
            'index_col': False,
            'row_header_score': 0.0,
            'col_header_score': 0.0,
            'row_headers': [],
            'col_headers': [],
        }

    # row header signals (first row)
    row_values = table_df.iloc[0].tolist()
    row_texts = [_cell_to_display_text(v, max_len=max_cell_chars, digit=digit).strip() for v in row_values]
    row_non_empty = [t for t in row_texts if t]
    row_non_empty_ratio = len(row_non_empty) / max(1, cols)
    row_text_like = [
        _is_text_like(row_values[i], row_texts[i]) for i in range(len(row_texts)) if row_texts[i]
    ]
    row_text_ratio = (sum(1 for x in row_text_like if x) / max(1, len(row_non_empty))) if row_non_empty else 0.0
    row_unique_ratio = len(set(t.lower() for t in row_non_empty)) / max(1, len(row_non_empty)) if row_non_empty else 0.0

    row_fill_ratio = 0.0
    if rows > 1 and row_non_empty:
        col_fill = []
        for cidx, header_text in enumerate(row_texts):
            if not header_text:
                continue
            col_values = table_df.iloc[1:, cidx].tolist()
            col_texts = [_cell_to_display_text(v, max_len=max_cell_chars, digit=digit).strip() for v in col_values]
            non_empty = sum(1 for t in col_texts if t)
            col_fill.append(non_empty / max(1, len(col_texts)))
        if col_fill:
            row_fill_ratio = sum(col_fill) / len(col_fill)

    row_header_score = round((row_non_empty_ratio + row_text_ratio + row_unique_ratio + row_fill_ratio) / 4, 3)
    index_row = row_header_score >= header_threshold

    # column header signals (first column)
    col_values = table_df.iloc[:, 0].tolist()
    col_texts = [_cell_to_display_text(v, max_len=max_cell_chars, digit=digit).strip() for v in col_values]
    col_non_empty = [t for t in col_texts if t]
    col_non_empty_ratio = len(col_non_empty) / max(1, rows)
    col_text_like = [
        _is_text_like(col_values[i], col_texts[i]) for i in range(len(col_texts)) if col_texts[i]
    ]
    col_text_ratio = (sum(1 for x in col_text_like if x) / max(1, len(col_non_empty))) if col_non_empty else 0.0
    col_unique_ratio = len(set(t.lower() for t in col_non_empty)) / max(1, len(col_non_empty)) if col_non_empty else 0.0

    col_fill_ratio = 0.0
    if cols > 1 and col_non_empty:
        row_fill = []
        for ridx, header_text in enumerate(col_texts):
            if not header_text:
                continue
            row_vals = table_df.iloc[ridx, 1:].tolist()
            row_texts_rest = [_cell_to_display_text(v, max_len=max_cell_chars, digit=digit).strip() for v in row_vals]
            non_empty = sum(1 for t in row_texts_rest if t)
            row_fill.append(non_empty / max(1, len(row_texts_rest)))
        if row_fill:
            col_fill_ratio = sum(row_fill) / len(row_fill)

    col_header_score = round((col_non_empty_ratio + col_text_ratio + col_unique_ratio + col_fill_ratio) / 4, 3)
    index_col = col_header_score >= header_threshold

    return {
        'index_row': index_row,
        'index_col': index_col,
        'row_header_score': row_header_score,
        'col_header_score': col_header_score,
        'row_headers': row_non_empty,
        'col_headers': col_non_empty,
    }


def _precompute_table_markdowns(
        segments: List[Dict[str, Any]],
        max_cell_chars: int,
        digit: int
    ) -> Dict[int, str]:
    """預先轉換所有 segments 的 DataFrame 為 Markdown 並快取"""
    markdown_map: Dict[int, str] = {}
    for idx, seg in enumerate(segments):
        table_df = seg.get('table_df')
        if table_df is None or getattr(table_df, 'empty', True):
            continue
        try:
            view_df = table_df.copy()
            view_df = view_df.applymap(lambda v: _cell_to_display_text(v, max_len=max_cell_chars, digit=digit))
            table_md = _df_to_markdown(view_df)
        except Exception:
            table_md = _df_to_markdown(table_df)
        markdown_map[idx] = table_md
    return markdown_map


def _select_table_index_by_llm(
        segments: List[Dict[str, Any]],
        max_cell_chars: int,
        digit: int,
        llm_provider: str,
        llm_model: str,
        llm_base_url: Optional[str],
        markdown_map: Optional[Dict[int, str]] = None,
        **kwargs
    ) -> Dict[int, Dict[str, Any]]:
    if not llm_base_url:
        llm_base_url = m_config.get('llm', {}).get('base_url', 'http://10.1.3.127:7017') if m_config else 'http://10.1.3.127:7017'

    # 如果沒有提供快取的 markdown_map，則預先計算
    if markdown_map is None:
        markdown_map = _precompute_table_markdowns(segments, max_cell_chars, digit)

    prompts: List[str] = []
    seg_indices: List[int] = []
    for idx, table_md in markdown_map.items():
        prompts.append(_build_llm_index_prompt(table_md))
        seg_indices.append(idx)

    if not prompts:
        return {}

    try:
        import requests
    except Exception:
        m_logger.warning("[chunk][index] 無法載入 requests，回退為空 index")
        return {}

    payload = {
        'prompts': prompts,
        'provider': llm_provider,
        'model': llm_model,
        'max_tokens': kwargs.get('llm_index_max_tokens', 32),
        'temperature': kwargs.get('llm_index_temperature', 0.0),
        'parallel': True,
        'max_batch_size': 190,
    }
    try:
        batch_chat_url = f"{llm_base_url.rstrip('/')}/chat/batch"
        resp = requests.post(batch_chat_url, json=payload, timeout=kwargs.get('llm_index_timeout', 120))
        resp.raise_for_status()
        data = resp.json()
        results = data.get('results', [])
    except Exception as e:
        m_logger.warning(f"[chunk][index] index 判斷失敗，回退為空: {e}")
        return {}

    index_map: Dict[int, Dict[str, Any]] = {}
    error_count = 0
    parse_fail_count = 0
    success_count = 0
    
    for i, item in enumerate(results):
        idx = seg_indices[i] if i < len(seg_indices) else None
        if idx is None:
            continue
        if item.get('error'):
            error_count += 1
            m_logger.warning(f"[chunk][index] seg_index={idx} LLM 錯誤: {item.get('error')}")
            continue
        result_obj = item.get('result', {}) if isinstance(item, dict) else {}
        output = ''
        if isinstance(result_obj, dict):
            output = result_obj.get('output', '') or ''
        parsed = _parse_index_from_llm_text(output)
        if parsed is None:
            parse_fail_count += 1
            m_logger.warning(f"[chunk][index] seg_index={idx} 解析失敗，LLM 輸出: {output[:1000]}+...")
            continue
        index_map[idx] = parsed
        success_count += 1
        m_logger.info(
            f"[chunk][index] seg_index={idx} index_row={parsed.get('index_row')} "
            f"index_col={parsed.get('index_col')} reason={parsed.get('reason')}"
        )
    
    if not index_map:
        m_logger.warning(
            f"[chunk][index] 所有 {len(seg_indices)} 個 segments 的 index 偵測失敗 "
            f"(成功={success_count}, 錯誤={error_count}, 解析失敗={parse_fail_count})"
        )
    else:
        m_logger.info(
            f"[chunk][index] index 偵測完成: 成功={success_count}/{len(seg_indices)}, "
            f"錯誤={error_count}, 解析失敗={parse_fail_count}"
        )
    
    return index_map


def _select_table_index_and_mode_by_llm(
        segments: List[Dict[str, Any]],
        max_cell_chars: int,
        digit: int,
        llm_provider: str,
        llm_model: str,
        llm_base_url: Optional[str],
        markdown_map: Optional[Dict[int, str]] = None,
        **kwargs
    ) -> Tuple[Dict[int, Dict[str, Any]], Dict[int, str]]:
    """合併的 LLM 調用：同時返回 index 和 mode 偵測結果"""
    if not llm_base_url:
        llm_base_url = m_config.get('llm', {}).get('base_url', 'http://10.1.3.127:7017') if m_config else 'http://10.1.3.127:7017'

    # 如果沒有提供快取的 markdown_map，則預先計算
    if markdown_map is None:
        markdown_map = _precompute_table_markdowns(segments, max_cell_chars, digit)

    prompts: List[str] = []
    seg_indices: List[int] = []
    for idx, table_md in markdown_map.items():
        table_df = segments[idx].get('table_df') if idx < len(segments) else None
        numeric_ratio = _compute_numeric_cell_ratio(table_df, max_cell_chars, digit) if table_df is not None else 0.0
        m_logger.info(f"[chunk][combined] seg_index={idx} numeric_ratio={numeric_ratio:.3f}")
        prompts.append(_build_llm_combined_prompt(table_md, numeric_ratio=numeric_ratio))
        seg_indices.append(idx)

    if not prompts:
        return {}, {}

    try:
        import requests
    except Exception:
        m_logger.warning("[chunk][combined] 無法載入 requests，回退為空結果")
        return {}, {}

    payload = {
        'prompts': prompts,
        'provider': llm_provider,
        'model': llm_model,
        'max_tokens': kwargs.get('llm_combined_max_tokens', 64),
        'temperature': kwargs.get('llm_combined_temperature', 0.0),
        'parallel': True,
        'max_batch_size': 190,
    }
    try:
        batch_chat_url = f"{llm_base_url.rstrip('/')}/chat/batch"
        resp = requests.post(batch_chat_url, json=payload, timeout=kwargs.get('llm_combined_timeout', 120))
        resp.raise_for_status()
        data = resp.json()
        results = data.get('results', [])
    except Exception as e:
        m_logger.warning(f"[chunk][combined] 合併偵測失敗，回退為空: {e}")
        return {}, {}

    index_map: Dict[int, Dict[str, Any]] = {}
    mode_map: Dict[int, str] = {}
    error_count = 0
    parse_fail_count = 0
    success_count = 0

    for i, item in enumerate(results):
        idx = seg_indices[i] if i < len(seg_indices) else None
        if idx is None:
            continue
        if item.get('error'):
            error_count += 1
            m_logger.warning(f"[chunk][combined] seg_index={idx} LLM 錯誤: {item.get('error')}")
            continue
        result_obj = item.get('result', {}) if isinstance(item, dict) else {}
        output = ''
        if isinstance(result_obj, dict):
            output = result_obj.get('output', '') or ''
        index_dict, mode = _parse_combined_from_llm_text(output)
        if index_dict is None and mode is None:
            parse_fail_count += 1
            m_logger.warning(f"[chunk][combined] seg_index={idx} 解析失敗，LLM 輸出: {output[:200]}")
            continue
        if index_dict:
            index_map[idx] = index_dict
        if mode:
            mode_map[idx] = mode
        else:
            mode_map[idx] = 'row'  # 預設回退
        success_count += 1
        m_logger.info(
            f"[chunk][combined] seg_index={idx} index_row={index_dict.get('index_row') if index_dict else None} "
            f"index_col={index_dict.get('index_col') if index_dict else None} mode={mode_map.get(idx)}"
        )

    if not index_map and not mode_map:
        m_logger.warning(
            f"[chunk][combined] 所有 {len(seg_indices)} 個 segments 的合併偵測失敗 "
            f"(成功={success_count}, 錯誤={error_count}, 解析失敗={parse_fail_count})"
        )
    else:
        m_logger.info(
            f"[chunk][combined] 合併偵測完成: 成功={success_count}/{len(seg_indices)}, "
            f"錯誤={error_count}, 解析失敗={parse_fail_count}"
        )

    return index_map, mode_map


def _select_table_modes_by_llm(
        segments: List[Dict[str, Any]],
        max_cell_chars: int,
        digit: int,
        llm_provider: str,
        llm_model: str,
        llm_base_url: Optional[str],
        index_signals_map: Optional[Dict[int, Dict[str, Any]]] = None,
        markdown_map: Optional[Dict[int, str]] = None,
        **kwargs
    ) -> Dict[int, str]:
    if not llm_base_url:
        llm_base_url = m_config.get('llm', {}).get('base_url', 'http://10.1.3.127:7017') if m_config else 'http://10.1.3.127:7017'

    # 如果沒有提供快取的 markdown_map，則預先計算
    if markdown_map is None:
        markdown_map = _precompute_table_markdowns(segments, max_cell_chars, digit)

    prompts: List[str] = []
    seg_indices: List[int] = []
    for idx, table_md in markdown_map.items():
        signals = index_signals_map.get(idx) if index_signals_map else None
        table_df = segments[idx].get('table_df') if idx < len(segments) else None
        numeric_ratio = _compute_numeric_cell_ratio(table_df, max_cell_chars, digit) if table_df is not None else 0.0
        m_logger.info(f"[chunk][llm] seg_index={idx} numeric_ratio={numeric_ratio:.3f}")
        prompts.append(_build_llm_table_prompt(table_md, signals, numeric_ratio))
        seg_indices.append(idx)

    if not prompts:
        return {}

    try:
        import requests
    except Exception:
        m_logger.warning("[chunk][llm] 無法載入 requests，回退為 row 模式")
        return {idx: 'row' for idx in seg_indices}

    payload = {
        'prompts': prompts,
        'provider': llm_provider,
        'model': llm_model,
        'max_tokens': kwargs.get('llm_mode_max_tokens', 8),
        'temperature': kwargs.get('llm_mode_temperature', 0.0),
        'parallel': True,
        'max_batch_size': 190,
    }
    try:
        batch_chat_url = f"{llm_base_url.rstrip('/')}/chat/batch"
        resp = requests.post(batch_chat_url, json=payload, timeout=kwargs.get('llm_mode_timeout', 120))
        resp.raise_for_status()
        data = resp.json()
        results = data.get('results', [])
    except Exception as e:
        m_logger.warning(f"[chunk][llm] 模式判斷失敗，回退為 row: {e}")
        return {idx: 'row' for idx in seg_indices}

    mode_map: Dict[int, str] = {}
    for i, item in enumerate(results):
        idx = seg_indices[i] if i < len(seg_indices) else None
        if idx is None:
            continue
        if item.get('error'):
            mode_map[idx] = 'row'
            continue
        result_obj = item.get('result', {}) if isinstance(item, dict) else {}
        output = ''
        if isinstance(result_obj, dict):
            output = result_obj.get('output', '') or ''
        mode, reason = _parse_mode_and_reason_from_llm_text(output)
        mode = mode or 'row'
        if mode not in {'row', 'column', 'preview', 'diagonal', 'cell'}:
            mode = 'row'
        mode_map[idx] = mode
        if reason:
            m_logger.info(f"[chunk][llm] seg_index={idx} mode={mode} reason={reason}")
        else:
            m_logger.info(f"[chunk][llm] seg_index={idx} mode={mode}")
    return mode_map


def chunk(
        segments: List[Dict[str, Any]],
        parse_mode: str = 'preview',
        max_cell_chars: int = 100,
        digit: int = 6,
        extract_kw_lbd: int = 10,
        llm_provider: str = 'remote',
        llm_model: str = 'remote8b',
        llm_base_url: str = None,
        metadata: Dict[str, Any] = None,
        use_llm_index: bool = True,
        **kwargs
    ) -> List[Dict[str, Any]]:
    if not segments:
        return []
    parse_mode = (parse_mode or 'row').lower().strip()
    if parse_mode not in {'row', 'column', 'cell', 'preview', 'diagonal', 'llm'}:
        raise ValueError('parse_mode 僅支援 row / column / cell / preview / diagonal / llm')

    chunks: List[Dict[str, Any]] = []
    chunk_order = 0
    # 批次處理：先跳過生成和去重，最後統一處理
    use_batch_dedup = dedup_multi_prompts_by_llm is not None
    use_batch_keyword_gen = parse_keywords_from_text is not None
    chunkers = {
        'row': RowChunker('row', max_cell_chars, digit, metadata, skip_dedup=use_batch_dedup, skip_keyword_gen=use_batch_keyword_gen),
        'column': ColumnChunker('column', max_cell_chars, digit, metadata, skip_dedup=use_batch_dedup, skip_keyword_gen=use_batch_keyword_gen),
        'cell': CellChunker('cell', max_cell_chars, digit, metadata, skip_dedup=use_batch_dedup, skip_keyword_gen=use_batch_keyword_gen),
        'preview': PreviewChunker('preview', max_cell_chars, digit, metadata, skip_dedup=use_batch_dedup, skip_keyword_gen=use_batch_keyword_gen),
        'diagonal': DiagonalChunker('diagonal', max_cell_chars, digit, metadata, skip_dedup=use_batch_dedup, skip_keyword_gen=use_batch_keyword_gen),
    }

    # 預先轉換並快取所有 segments 的 Markdown（避免重複轉換）
    markdown_map: Optional[Dict[int, str]] = None
    if parse_mode in {'row', 'column', 'llm'} and (use_llm_index or parse_mode == 'llm'):
        markdown_map = _precompute_table_markdowns(segments, max_cell_chars, digit)

    llm_mode_map: Dict[int, str] = {}
    index_signals_map: Dict[int, Dict[str, Any]] = {}
    
    # 當 parse_mode='llm' 且 use_llm_index=True 時，使用合併的 LLM 調用
    if parse_mode == 'llm' and use_llm_index:
        index_signals_map, llm_mode_map = _select_table_index_and_mode_by_llm(
            segments=segments,
            max_cell_chars=max_cell_chars,
            digit=digit,
            llm_provider=llm_provider,
            llm_model=llm_model,
            llm_base_url=llm_base_url,
            markdown_map=markdown_map,
            **kwargs
        )
        m_logger.info(_log_prefix('OKBLUE') + f"[chunk] 合併偵測完成: index_signals_map={len(index_signals_map)}, mode_map={len(llm_mode_map)}")
    else:
        # 分別調用（向後兼容）
        if parse_mode in {'row', 'column', 'llm'} and use_llm_index:
            index_signals_map = _select_table_index_by_llm(
                segments=segments,
                max_cell_chars=max_cell_chars,
                digit=digit,
                llm_provider=llm_provider,
                llm_model=llm_model,
                llm_base_url=llm_base_url,
                markdown_map=markdown_map,
                **kwargs
            )
            m_logger.info(_log_prefix('OKBLUE') + f"[chunk] use_llm_index={use_llm_index} index_signals_map={index_signals_map}")
        if parse_mode == 'llm':
            llm_mode_map = _select_table_modes_by_llm(
                segments=segments,
                max_cell_chars=max_cell_chars,
                digit=digit,
                llm_provider=llm_provider,
                llm_model=llm_model,
                llm_base_url=llm_base_url,
                index_signals_map=index_signals_map,
                markdown_map=markdown_map,
                **kwargs
            )
    # 先收集所有 chunks（不進行去重）
    for idx, seg in enumerate(segments):
        table_df = seg.get('table_df')
        if table_df is None or getattr(table_df, 'empty', True):
            chunks.append(_build_chunk_from_text(
                chunk_text=seg.get('unit_text', ''),
                seg=seg,
                parse_mode=parse_mode,
                chunk_order=chunk_order,
                metadata=metadata,
                row_index=None,
                col_index=None,
                cell_ref=None,
                skip_dedup=use_batch_dedup,  # 先跳過去重
                skip_keyword_gen=use_batch_keyword_gen,  # 先跳過生成
            ))
            chunk_order += 1
            continue

        if parse_mode == 'llm':
            mode = llm_mode_map.get(idx, 'row')
            numeric_ratio = _compute_numeric_cell_ratio(table_df, max_cell_chars, digit)
            if numeric_ratio > 0.5:
                mode = 'preview'
                try:
                    m_logger.info(f"[chunk][llm] seg_index={idx} numeric_ratio={numeric_ratio:.3f} force_mode=preview")
                except Exception:
                    pass
        else:
            mode = parse_mode
        try:
            m_logger.info(f"[chunk] seg_index={idx} parse_mode={parse_mode} use_mode={mode}")
        except Exception:
            pass
        # 預先計算並快取 headers（避免重複計算）
        cached_headers = _guess_headers(table_df)
        chunker = chunkers.get(mode, chunkers['row'])
        seg_chunks, chunk_order = chunker.generate(
            seg, table_df, chunk_order, index_signals=index_signals_map.get(idx), cached_headers=cached_headers
        )
        if seg_chunks:
            chunks.extend(seg_chunks)
    
    # 批次生成 keywords：收集所有 chunk_text 後批次處理
    if parse_keywords_from_text is not None and chunks:
        try:
            # 收集需要生成 keywords 的 chunks（長度 >= 10）
            chunk_texts: List[str] = []
            chunk_indices_for_gen: List[int] = []
            for i, ch in enumerate(chunks):
                chunk_text = ch.get('text', '') or ch.get('unit_text', '')
                if chunk_text and len(chunk_text) >= 10:
                    chunk_texts.append(chunk_text)
                    chunk_indices_for_gen.append(i)
            
            # 批次生成 keywords（如果 parse_keywords_from_text 支持批次）
            # 注意：目前 parse_keywords_from_text 是單文本處理，需要批次化包裝
            if chunk_texts:
                batch_size = kwargs.get('keywords_batch_size', 50)  # 批次大小限制
                generated_keywords: List[List[str]] = []
                
                # 分批處理，避免一次處理過多
                for start in range(0, len(chunk_texts), batch_size):
                    batch_texts = chunk_texts[start:start + batch_size]
                    batch_indices = chunk_indices_for_gen[start:start + batch_size]
                    
                    # 對每個文本生成 keywords（目前是單個處理，未來可優化為真正的批次）
                    batch_keywords = []
                    for text in batch_texts:
                        try:
                            kw = parse_keywords_from_text(text) or []
                            batch_keywords.append(kw)
                        except Exception:
                            batch_keywords.append([])
                    
                    generated_keywords.extend(batch_keywords)
                
                # 將生成的 keywords 分配回對應的 chunks
                for idx, keywords in zip(chunk_indices_for_gen, generated_keywords):
                    if keywords:
                        # 過濾無意義標籤
                        if filter_meaningless_tags is not None:
                            try:
                                keywords = filter_meaningless_tags(keywords) or keywords
                            except Exception:
                                pass
                        chunks[idx]['multi_prompts'] = keywords
                
                m_logger.info(f"[chunk] 批次生成 keywords 完成: {len(chunk_indices_for_gen)} 個 chunks")
        except Exception as e:
            m_logger.warning(f"[chunk] 批次生成 keywords 失敗，回退為單個處理: {e}")
            # 回退：對每個 chunk 單獨生成
            for ch in chunks:
                chunk_text = ch.get('text', '') or ch.get('unit_text', '')
                if chunk_text and len(chunk_text) >= 10:
                    try:
                        keywords = parse_keywords_from_text(chunk_text) or []
                        if filter_meaningless_tags is not None:
                            try:
                                keywords = filter_meaningless_tags(keywords) or keywords
                            except Exception:
                                pass
                        ch['multi_prompts'] = keywords
                    except Exception:
                        pass
    
    # 批次去重：收集所有 keywords 後批次處理
    if use_batch_dedup and chunks:
        try:
            # 收集所有 keywords 列表
            all_keywords_lists: List[List[str]] = []
            chunk_indices: List[int] = []  # 記錄每個 keywords 對應的 chunk index
            for i, ch in enumerate(chunks):
                keywords = ch.get('multi_prompts', [])
                if keywords and isinstance(keywords, list):
                    all_keywords_lists.append(keywords)
                    chunk_indices.append(i)
            
            # 批次去重（dedup_multi_prompts_by_llm 內部有 batch_size 限制，預設 8）
            if all_keywords_lists:
                deduped_lists = dedup_multi_prompts_by_llm(all_keywords_lists)
                if deduped_lists and len(deduped_lists) == len(all_keywords_lists):
                    # 將去重後的結果分配回對應的 chunks
                    for idx, deduped_kw in zip(chunk_indices, deduped_lists):
                        if isinstance(deduped_kw, list):
                            chunks[idx]['multi_prompts'] = deduped_kw
                        elif deduped_kw:
                            # 如果返回的不是 list，嘗試轉換
                            chunks[idx]['multi_prompts'] = [str(deduped_kw)] if deduped_kw else []
                    m_logger.info(f"[chunk] 批次去重完成: {len(all_keywords_lists)} 個 keywords 列表")
        except Exception as e:
            m_logger.warning(f"[chunk] 批次去重失敗，回退為單個處理: {e}")
            # 回退：對每個 chunk 單獨去重
            for ch in chunks:
                keywords = ch.get('multi_prompts', [])
                if keywords and isinstance(keywords, list):
                    try:
                        deduped = dedup_multi_prompts_by_llm([keywords])
                        if deduped and isinstance(deduped, list) and len(deduped) > 0:
                            ch['multi_prompts'] = deduped[0] if isinstance(deduped[0], list) else keywords
                    except Exception:
                        pass
    
    return chunks



def _build_chunk_from_text(
        chunk_text: str,
        seg: Dict[str, Any],
        parse_mode: str,
        chunk_order: int,
        metadata: Optional[Dict[str, Any]],
        row_index: Optional[int],
        col_index: Optional[int],
        cell_ref: Optional[str],
        skip_dedup: bool = False,
        skip_keyword_gen: bool = False,
    ) -> Dict[str, Any]:
    keywords: List[str] = []
    # 如果跳過生成，使用空列表（將在批次處理中生成）
    if not skip_keyword_gen:
        if parse_keywords_from_text is not None and len(chunk_text) >= 10:
            try:
                keywords = parse_keywords_from_text(chunk_text) or []
            except Exception:
                keywords = []
        if not keywords:
            keywords = _simple_keywords(chunk_text)
        if filter_meaningless_tags is not None:
            try:
                keywords = filter_meaningless_tags(keywords)
            except Exception:
                pass
        if not skip_dedup and dedup_multi_prompts_by_llm is not None:
            try:
                keywords = dedup_multi_prompts_by_llm([keywords]) or keywords
            except Exception:
                pass
    
    # 確保 keywords 是 List[str]，扁平化嵌套列表
    flattened_keywords: List[str] = []
    for kw in keywords:
        if isinstance(kw, str):
            flattened_keywords.append(kw)
        elif isinstance(kw, (list, tuple)):
            flattened_keywords.extend(str(item) for item in kw if item)
        else:
            flattened_keywords.append(str(kw))
    keywords = flattened_keywords

    call_prompt = chunk_text
    return {
        'chunk_id': f'chunk_{chunk_order:04d}',
        'order': chunk_order,
        'unit_text': chunk_text,
        'text': chunk_text,
        'call_prompt': call_prompt,
        'multi_prompts': keywords,
        'indent_level': 0,
        'parse_mode': parse_mode,
        'sheet_name': seg.get('sheet_name'),
        'sheet_index': seg.get('sheet_index'),
        'segment_id': seg.get('segment_id'),
        'segment_type': seg.get('segment_type'),
        'table_id': seg.get('table_id'),
        'row_index': row_index,
        'column_index': col_index,
        'cell_ref': cell_ref,
        'table_bbox': [seg.get('top_row'), seg.get('left_col'), seg.get('bottom_row'), seg.get('right_col')],
        'headers': seg.get('headers', []),
        'images': seg.get('images', []),
        'sheet_images': seg.get('sheet_images', []),
        'formula_cells': seg.get('formula_cells', []),
        'metadata': {
            'sheet_name': seg.get('sheet_name'),
            'table_id': seg.get('table_id'),
            'parse_mode': parse_mode,
        }
    }



def _simple_keywords(text: str, max_keywords: int = 10) -> List[str]:
    tokens = re.findall(r'[A-Za-z0-9_\-\u4e00-\u9fff]{2,}', text)
    seen = set()
    results = []
    for token in tokens:
        t = token.strip()
        if not t:
            continue
        key = t.lower()
        if key in seen:
            continue
        seen.add(key)
        results.append(t)
        if len(results) >= max_keywords:
            break
    return results


# ============================================================
# 5. Process（一條龍）
# ============================================================

def process(
    file_path: str,
    parse_mode: str = 'preview',
        max_cell_chars: int = 100,
        digit: int = 6,
        suitable_char_count: int = 500,
        include_images: bool = True,
        image_placeholder: bool = True,
        enable_image_llm: bool = True,
        llm_provider: str = 'remote',
        llm_model: str = 'remote8b',
        llm_base_url: str = None,
        image_llm_provider: str = 'openai',
        image_llm_model: str = 'gpt4o_chat',
        **kwargs
    ) -> Dict[str, Any]:
    m_logger.info(f'[process] 開始處理檔案: {file_path}')
    extract_result = extract(
        file_path=file_path,
        include_images=include_images,
        image_placeholder=image_placeholder,
        **kwargs
    )
    unit_paras = extract_result['unit_paras']
    metadata = extract_result.get('metadata', {})

    segments = segment(
        unit_paras=unit_paras,
        suitable_char_count=suitable_char_count,
        metadata=metadata,
        enable_image_llm=enable_image_llm,
        llm_provider=llm_provider,
        llm_model=llm_model,
        llm_base_url=llm_base_url,
        image_llm_provider=image_llm_provider,
        image_llm_model=image_llm_model,
        **kwargs
    )
    chunks = chunk(
        segments=segments,
        parse_mode=parse_mode,
        max_cell_chars=max_cell_chars,
        digit=digit,
        llm_provider=llm_provider,
        llm_model=llm_model,
        llm_base_url=llm_base_url,
        metadata=metadata,
        **kwargs
    )

    stats = {
        'sheet_count': metadata.get('sheet_count', 0),
        'sheet_units': len(unit_paras),
        'segment_count': len(segments),
        'chunk_count': len(chunks),
        'table_count': metadata.get('total_tables', 0),
        'formula_count': metadata.get('formula_count', 0),
        'image_count': metadata.get('image_count', 0),
        'parse_mode': parse_mode,
    }

    return {
        'chunks': chunks,
        'unit_paras': unit_paras,
        'segments': segments,
        'metadata': metadata,
        'metadata_display': _sanitize_metadata(metadata),
        'stats': stats,
    }


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python xlsx_parser.py <excel_file_path> [row|column|cell]')
        sys.exit(1)

    file_path = sys.argv[1]
    parse_mode = sys.argv[2] if len(sys.argv) > 2 else 'row'

    print('=== Preview ===')
    print(preview(file_path)[:2000])

    print('\n=== Process ===')
    result = process(file_path, parse_mode=parse_mode, enable_image_llm=False)
    print(json.dumps(result['stats'], ensure_ascii=False, indent=2))
    print('\n=== Metadata Summary ===')
    print(json.dumps(result['metadata_display'], ensure_ascii=False, indent=2))
