#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import logging
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    from ContextParser import xlsx_parser
except Exception:
    try:
        from .ContextParser import xlsx_parser  # type: ignore
    except Exception:
        xlsx_parser = None


def _setup_logger() -> logging.Logger:
    logger = logging.getLogger('file_importers')
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
        logger.addHandler(handler)
    return logger


m_logger = _setup_logger()

m_type_map = {
    '.xlsx': 'xlsx', '.xls': 'xlsx', '.xlsm': 'xlsx', '.csv': 'xlsx',
    '.txt': 'txt', '.text': 'txt', '.md': 'md', '.markdown': 'md',
    '.json': 'json',
}


def normalize_source_to_uri(source_input: str, document_id: Optional[str] = None) -> str:
    source_input = (source_input or '').strip()
    if not source_input:
        return f'doc://{document_id}' if document_id else ''
    if '://' in source_input:
        return source_input
    abs_path = os.path.abspath(os.path.expanduser(source_input))
    return f'file://{abs_path.replace(os.sep, "/")}'


class EncodingHandler:
    def __init__(self, logger: Optional[logging.Logger] = None):
        self.logger = logger or m_logger

    def decode_auto(self, raw_data: bytes) -> str:
        for enc in ('utf-8', 'utf-8-sig', 'cp950', 'big5', 'latin-1'):
            try:
                return raw_data.decode(enc)
            except Exception:
                continue
        return raw_data.decode('utf-8', errors='ignore')


def decode_file(file_path: str):
    try:
        with open(file_path, 'rb') as f:
            raw = f.read()
        return EncodingHandler().decode_auto(raw), 'auto_detected', None
    except Exception as e:
        return None, None, e


@dataclass
class DocumentPathResolver:
    project_root: str = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    def resolve(
        self,
        document_path: Optional[str],
        document_source_uri: Optional[str] = None,
        document_id: Optional[str] = None,
    ) -> Dict[str, Optional[str]]:
        source_input = document_source_uri or document_path or ''
        payload_source = normalize_source_to_uri(source_input, document_id)
        abs_path = os.path.abspath(document_path) if document_path else None
        return {
            'abs_path': abs_path,
            'document_source_uri': payload_source,
            'payload_source': payload_source,
            'source_key': payload_source,
        }


class FileImportManager:
    def __init__(self):
        self.logger = m_logger

    def get_parser_type(self, file_ext: str) -> str:
        return m_type_map.get(file_ext.lower(), file_ext.lstrip('.').lower())

    def get_parser_module(self, file_ext: str):
        parser_type = self.get_parser_type(file_ext)
        if parser_type == 'xlsx':
            return xlsx_parser
        return None

    def import_file(self, file_path: str, file_type: Optional[str] = None, **kwargs) -> Dict[str, Any]:
        file_ext = Path(file_path).suffix.lower()
        parser_type = file_type or self.get_parser_type(file_ext)

        if parser_type == 'xlsx':
            if xlsx_parser is None:
                return {'success': False, 'error': 'xlsx_parser not available', 'parser_type': 'xlsx'}

            preview_text = ''
            try:
                preview_text = xlsx_parser.preview(file_path)
            except Exception as e:
                self.logger.warning(f'preview failed: {e}')

            try:
                extract_result = xlsx_parser.extract(file_path=file_path, **kwargs)
            except Exception as e:
                return {'success': False, 'error': str(e), 'parser_type': 'xlsx'}

            return {
                'success': True,
                'parser_type': 'xlsx',
                'file_path': file_path,
                'text': preview_text,
                **extract_result,
            }

        if parser_type in {'txt', 'md'}:
            text, used_encoding, err = decode_file(file_path)
            if err:
                return {'success': False, 'error': str(err), 'parser_type': parser_type}
            return {
                'success': True,
                'parser_type': parser_type,
                'file_path': file_path,
                'text': text or '',
                'used_encoding': used_encoding,
            }

        if parser_type == 'json':
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    obj = json.load(f)
            except Exception as e:
                return {'success': False, 'error': str(e), 'parser_type': 'json'}
            text = json.dumps(obj, ensure_ascii=False)
            return {
                'success': True,
                'parser_type': 'json',
                'file_path': file_path,
                'text': text,
                'payloads': [{'type': 'json', 'text': text, 'source': file_path, 'data': obj}],
            }

        return {'success': False, 'error': f'unsupported parser type: {parser_type}', 'parser_type': parser_type}

    def process_file(self, ret: Dict[str, Any], local_path: Optional[str] = None, filename: Optional[str] = None, **kwargs) -> bool:
        file_path = local_path or ret.get('file_path')
        if not file_path:
            ret.update({'success': False, 'error': 'missing file_path'})
            return False

        parser_type = self.get_parser_type(Path(file_path).suffix)
        if parser_type == 'xlsx':
            if xlsx_parser is None or not hasattr(xlsx_parser, 'process'):
                ret.update({'success': False, 'error': 'xlsx process parser unavailable'})
                return False
            try:
                process_result = xlsx_parser.process(file_path=file_path, **kwargs)
            except Exception as e:
                ret.update({'success': False, 'error': str(e)})
                return False
            ret.update({'success': True, 'parser_type': 'xlsx', 'file_path': file_path, **process_result})
            return True

        text = ret.get('text', '')
        segment = {'content': text, 'metadata': {'sequence': 1, 'type': parser_type}}
        chunk = {'chunk_id': 'chunk_0001', 'order': 0, 'call_prompt': text, 'text': text}
        ret.update({
            'success': True,
            'parser_type': parser_type,
            'file_path': file_path,
            'unit_paras': [segment],
            'segments': [segment],
            'chunks': [chunk],
            'stats': {'segment_count': 1, 'chunk_count': 1, 'parse_mode': kwargs.get('parse_mode', 'preview')},
        })
        return True

    def run_stage_extract(self, local_path: str, filename: str = None, config: Dict[str, Any] = None) -> Dict[str, Any]:
        return self.import_file(local_path)

    def run_stage_segment(self, import_result: Dict[str, Any], **kwargs) -> Dict[str, Any]:
        if import_result.get('parser_type') == 'xlsx' and xlsx_parser is not None:
            unit_paras = import_result.get('unit_paras', [])
            metadata = import_result.get('metadata', {})
            try:
                segments = xlsx_parser.segment(unit_paras=unit_paras, metadata=metadata, **kwargs)
                return {'success': True, 'segments': segments}
            except Exception as e:
                return {'success': False, 'error': str(e)}
        text = import_result.get('text', '')
        return {'success': True, 'segments': [{'content': text, 'metadata': {'sequence': 1}}]}

    def run_stage_chunk(self, segment_result: Dict[str, Any], **kwargs) -> Dict[str, Any]:
        segments = segment_result.get('segments', [])
        if xlsx_parser is not None and segments and isinstance(segments[0], dict) and 'table_df' in segments[0]:
            try:
                chunks = xlsx_parser.chunk(segments=segments, **kwargs)
                return {'success': True, 'chunks': chunks}
            except Exception as e:
                return {'success': False, 'error': str(e)}
        text = ''
        if segments:
            text = segments[0].get('content', '')
        return {'success': True, 'chunks': [{'chunk_id': 'chunk_0001', 'order': 0, 'call_prompt': text, 'text': text}]}


def detect_file_change(*, path: str, prev_snapshot: Optional[Dict[str, Any]] = None) -> bool:
    if not path or not os.path.exists(path):
        return False
    st = os.stat(path)
    curr = {'mtime': int(st.st_mtime), 'size': int(st.st_size)}
    if not prev_snapshot:
        return True
    return prev_snapshot.get('mtime') != curr['mtime'] or prev_snapshot.get('size') != curr['size']


def detect_folder_change(*, path: str, prev_snapshot: Optional[Dict[str, Any]] = None, recursive: bool = True, include: Optional[List[str]] = None, exclude: Optional[List[str]] = None) -> bool:
    if not path or not os.path.exists(path):
        return False
    curr = {'mtime': int(os.path.getmtime(path))}
    if not prev_snapshot:
        return True
    return prev_snapshot.get('mtime') != curr['mtime']


if __name__ == '__main__':
    manager = FileImportManager()
    print('file_importers local module ready')
