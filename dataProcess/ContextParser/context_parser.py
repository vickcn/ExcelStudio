#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ContextParser - 文件書寫模式檢測與分段策略

用於偵測文件書寫模式（code/log/communication/structured_doc）
並選擇對應的分段策略，生成符合系統定義的 chunk
"""

import os
import json
import re
import logging
import uuid
import math
from typing import List, Dict, Any, Optional, Tuple, Union, Set
from pathlib import Path
import queue, threading, time
import requests
from abc import ABC, abstractmethod

import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from package import LOGger
from package import dataframeprocedure as DFP
from src.parallel_llm_executor import ParallelLLMExecutor
from src.image_explainer_output import finalize_image_explainer_text


# 嘗試導入 chardet，如果沒有則使用備用方案
try:
    import chardet
    HAS_CHARDET = True
except ImportError:
    HAS_CHARDET = False

def _setup_logger() -> logging.Logger:
        """設定日誌"""
        # 檔案處理器
        log_dir = os.path.join(os.path.dirname(__file__), 'log')
        os.makedirs(log_dir, exist_ok=True)
        
        logger = LOGger.addloger(logfile=os.path.join(log_dir, 'context_parser_%t.log'))
        logger.error = lambda x,*args,colora=LOGger.FAIL,**kwargs: logger(x,*args,**kwargs, colora=colora)
        logger.warning = lambda x,*args,colora=LOGger.WARNING,**kwargs: logger(x,*args,**kwargs, colora=colora)
        logger.info = lambda x,*args,**kwargs: logger(x,*args,**kwargs)
        logger.summary = lambda x,*args,colora=LOGger.OKCYAN,**kwargs: logger(x,*args,**kwargs, colora=colora)
        logger.debug = lambda x,*args,colora=LOGger.OKBLUE,**kwargs: logger(x,*args,**kwargs, colora=colora)
        logger.exception = lambda x,logfile='',**kwargs: LOGger.exception_process(x,logfile=logfile,**kwargs)
        return logger

m_logger = _setup_logger()
m_fn = os.path.basename(__file__).replace('.py', '')
m_debug = LOGger.myDebuger(stamps=[m_fn])
m_config_file = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'config.json')
m_config = LOGger.load_json(m_config_file)
m_llm_error_patterns = set(m_config.get('llm_error', {}).get('patterns', []))

import unicodedata

def analyze_line_break_pattern(text: str, sample_size: int = 500) -> str:
    """分析文本前 m 個字符的換行符使用習慣"""
    if not text or len(text) < 10:
        return '\n\n'  # 預設使用雙換行
    
    sample = text[:min(sample_size, len(text))]
    
    # 統計各種換行符組合（優先檢查段落分隔符）
    patterns = {
        '\r\n\r\n': sample.count('\r\n\r\n'),  # Windows 雙換行（段落分隔）
        '\n\n': sample.count('\n\n'),          # Unix 雙換行
        '\r\n': sample.count('\r\n'),          # Windows 單換行
        '\n': sample.count('\n'),              # Unix 單換行
    }
    
    m_logger.debug(f"換行符分析結果: {patterns}")
    
    # 判斷主要使用的段落分隔符（優先雙換行）
    crlf_double_count = patterns['\r\n\r\n']
    lf_double_count = patterns['\n\n']
    crlf_single_count = patterns['\r\n']
    lf_single_count = patterns['\n']
    
    if crlf_double_count > 0:
        m_logger.info(f"檢測到 Windows 雙換行模式 (\\r\\n\\r\\n)，出現 {crlf_double_count} 次")
        return '\r\n\r\n'
    elif lf_double_count > 0:
        m_logger.info(f"檢測到 Unix 雙換行模式 (\\n\\n)，出現 {lf_double_count} 次")
        return '\n\n'
    elif crlf_single_count > lf_single_count * 2:
        # Windows 單換行明顯多於 Unix 單換行
        m_logger.info(f"檢測到 Windows 單換行模式 (\\r\\n)，出現 {crlf_single_count} 次")
        return '\r\n'
    else:
        # 預設使用 Unix 單換行
        m_logger.info(f"使用預設 Unix 單換行模式 (\\n)，出現 {lf_single_count} 次")
        return '\n'


def clean_rtf_remnants(text: str) -> str:
    """
    移除 RTF/HTML 轉文字後的殘留控制序列，供 msg_parser、html_parser 等共用。
    處理：字型表、\\*、&nbsp;、behavior:url、HYPERLINK、多餘空白等。
    """
    if not text or not isinstance(text, str):
        return text
    t = text
    # Strip Exchange/RTF converter headers like:
    # "微軟正黑體;新細明體; ... Microsoft Exchange Server; Converter converted from html; BM_BEGIN d"
    t = re.sub(
        r'^\s*[^\r\n]*?\bConverter\s+converted\s+from\s+html;?\s*BM_BEGIN\b\s*d\s*',
        '',
        t,
        flags=re.IGNORECASE
    )
    t = re.sub(r'^(?:(?:[A-Za-z][A-Za-z0-9 ]*|\"[^\"]*\")\s*;+\s*)+', '', t)
    t = re.sub(r'\\\s*\*', ' ', t)
    t = re.sub(r'&nbsp;', ' ', t)
    t = re.sub(r'\\[a-z]+\d*\\:*\s*', '', t)
    t = re.sub(r'\\behavior:url\([^)]*\)', '', t)
    t = re.sub(r'\\.shape\s*\\behavior[^;]*;?', '', t)
    t = re.sub(r'HYPERLINK\s+"[^"]*"\s*', '', t)
    # Preserve line breaks while normalizing whitespace inside lines.
    t = re.sub(r'\r\n?', '\n', t)
    # Drop leading Exchange converter header block if present.
    lines = t.split('\n')
    bm_idx = next((i for i, line in enumerate(lines[:20]) if 'BM_BEGIN' in line), None)
    if bm_idx is not None:
        start = bm_idx + 1
        while start < len(lines) and lines[start].strip().lower() in ('', 'd'):
            start += 1
        lines = lines[start:]
    # Drop stray single-letter "d" lines (RTF noise).
    lines = [line for line in lines if line.strip().lower() != 'd']
    # Remove leading "d" tokens before CJK text (RTF noise like "d 如果..." or "d待討論").
    lines = [re.sub(r'^\s*d\s*(?=[\u4e00-\u9fff])', '', line, flags=re.IGNORECASE) for line in lines]
    t = '\n'.join(lines)
    t = re.sub(r'[ \t]+', ' ', t)
    t = re.sub(r'\n{3,}', '\n\n', t)
    return t.strip()


def safe_filename(filename: str) -> str:
    # 去除左右空白
    filename = filename.strip()

    # 轉成 NFC，避免怪異組合字元
    filename = unicodedata.normalize("NFC", filename)

    # 移除不合法字元： / : * ? " < > | 及控制字元
    filename = re.sub(r'[/*?"<>|\x00-\x1f]', "‧", filename)

    # 避免檔名過長
    basename = os.path.basename(filename)[:200]
    dirname = os.path.dirname(filename)
    # 避免空字串
    return os.path.join(dirname, basename) or "unnamed"

def parse_keywords_from_text(keywords_text: str, max_keywords: int = 5, min_keyword_length: int = 1) -> List[str]:
    """
    從文本中解析關鍵詞，支持多種格式：
    1. JSON 格式：{"keywords": [...]} 或 [...]
    2. Markdown 代碼塊中的 JSON
    3. 第一個完整的 JSON 陣列或物件
    4. 逗號/分號分隔的文本
    
    Args:
        keywords_text: 包含關鍵詞的文本
        max_keywords: 最大關鍵詞數量
        min_keyword_length: 最小關鍵詞長度（默認為1，即長度必須 > 1，過濾單字符）
        
    Returns:
        關鍵詞列表
    """
    if not keywords_text or not isinstance(keywords_text, str):
        return []
    
    keywords = []
    keywords_text_stripped = keywords_text.strip()
    normalized = keywords_text_stripped.replace('，', ',').replace('：', ':').replace('；', ';')
    
    # 嘗試 1: 解析直接 JSON 格式（{"keywords": [...]} 或 [...]）
    candidates = [keywords_text_stripped, normalized]
    for candidate in candidates:
        if (candidate.startswith('{') and candidate.endswith('}')) or \
           (candidate.startswith('[') and candidate.endswith(']')):
            try:
                keywords_data = json.loads(candidate)
                if isinstance(keywords_data, dict):
                    keywords = keywords_data.get('keywords', [])
                elif isinstance(keywords_data, list):
                    keywords = keywords_data
                # 確保所有元素都是字符串，並過濾長度
                if isinstance(keywords, list):
                    keywords = [str(k).strip() for k in keywords if k and len(str(k).strip()) > min_keyword_length]
                    if keywords:
                        return keywords[:max_keywords]


            except (json.JSONDecodeError, ValueError, TypeError, AttributeError):
                pass
    
    # 嘗試 2: 提取 Markdown 代碼塊中的 JSON
    for candidate in [keywords_text_stripped, normalized]:
        # 匹配陣列格式
        code_block_match = re.search(r'```(?:json)?\s*(\[.*?\])\s*```', candidate, re.DOTALL)
        if not code_block_match:
            # 匹配物件格式（兼容舊格式）
            code_block_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', candidate, re.DOTALL)
        if code_block_match:
            try:
                keywords_data = json.loads(code_block_match.group(1))
                if isinstance(keywords_data, dict):
                    keywords = keywords_data.get('keywords', [])
                elif isinstance(keywords_data, list):
                    keywords = keywords_data
                if isinstance(keywords, list):
                    keywords = [str(k).strip() for k in keywords if k and len(str(k).strip()) > min_keyword_length]
                    if keywords:
                        return keywords[:max_keywords]
            except (json.JSONDecodeError, ValueError, TypeError, AttributeError):
                pass
    
    # 嘗試 3: 提取第一個完整的 JSON 陣列或物件
    for candidate in [keywords_text_stripped, normalized]:
        # 先嘗試陣列格式（使用配對括號查找）
        list_start_idx = candidate.find('[')
        if list_start_idx != -1:
            bracket_count = 0
            list_end_idx = -1
            for i in range(list_start_idx, len(candidate)):
                if candidate[i] == '[':
                    bracket_count += 1
                elif candidate[i] == ']':
                    bracket_count -= 1
                    if bracket_count == 0:
                        list_end_idx = i
                        break
            
            if list_end_idx != -1 and list_end_idx > list_start_idx:
                try:
                    json_str = candidate[list_start_idx:list_end_idx + 1]
                    keywords_data = json.loads(json_str)
                    if isinstance(keywords_data, list):
                        keywords = [str(k).strip() for k in keywords_data if k and len(str(k).strip()) > min_keyword_length]
                        if keywords:
                            return keywords[:max_keywords]
                except (json.JSONDecodeError, ValueError, TypeError, AttributeError):
                    pass
        
        # 再嘗試物件格式（兼容舊格式）
        obj_start_idx = candidate.find('{')
        if obj_start_idx != -1:
            bracket_count = 0
            obj_end_idx = -1
            for i in range(obj_start_idx, len(candidate)):
                if candidate[i] == '{':
                    bracket_count += 1
                elif candidate[i] == '}':
                    bracket_count -= 1
                    if bracket_count == 0:
                        obj_end_idx = i
                        break
            
            if obj_end_idx != -1 and obj_end_idx > obj_start_idx:
                try:
                    json_str = candidate[obj_start_idx:obj_end_idx + 1]
                    keywords_data = json.loads(json_str)
                    if isinstance(keywords_data, dict):
                        keywords = keywords_data.get('keywords', [])
                    if isinstance(keywords, list):
                        keywords = [str(k).strip() for k in keywords if k and len(str(k).strip()) > min_keyword_length]
                        if keywords:
                            return keywords[:max_keywords]
                except (json.JSONDecodeError, ValueError, TypeError, AttributeError):
                    pass
    
    # 嘗試 4: 處理不完整的 JSON 片段（例如：`["資訊部門"` 或 `["關鍵詞1", "關鍵詞2"`）
    # 這種情況通常發生在 LLM 回應被截斷或清洗後留下不完整的 JSON
    if keywords_text_stripped.startswith('[') or keywords_text_stripped.startswith('["'):
        # 先移除開頭的 `[` 和可能的 `"`
        text_to_process = keywords_text_stripped.lstrip('[').strip()
        # 移除轉義引號
        text_to_process = text_to_process.replace('\\"', '"')
        
        # 嘗試提取所有引號內的內容
        quoted_keywords = re.findall(r'"([^"]+)"', text_to_process)
        if quoted_keywords:
            keywords = [kw.strip() for kw in quoted_keywords if kw.strip() and len(kw.strip()) > min_keyword_length]
            if keywords:
                return keywords[:max_keywords]
        
        # 如果沒有找到引號內的內容，嘗試提取第一個引號後的內容（直到下一個引號或結尾）
        # 例如：`["資訊部門"` → `資訊部門`
        first_quote_match = re.search(r'^["\']?([^"\',\]]+)', text_to_process)
        if first_quote_match:
            potential_keyword = first_quote_match.group(1).strip()
            if potential_keyword and len(potential_keyword) > min_keyword_length:
                return [potential_keyword]
    
    # 嘗試 5: 如果文本看起來像 JSON 格式但解析失敗，返回空列表
    if keywords_text_stripped.startswith('[') or keywords_text_stripped.startswith('{'):
        return []
    
    # 嘗試 6: 使用分隔符分割（僅當不是 JSON 格式時）
    separators = [',', ';', '\n', '、', '||']  # 添加 || 分隔符支持
    for sep in separators:
        if sep in keywords_text_stripped:
            keywords = [kw.strip() for kw in keywords_text_stripped.split(sep) if kw.strip()]
            # 過濾掉看起來像 JSON 片段的部分，並應用最小長度過濾
            keywords = [kw for kw in keywords 
                      if len(kw) > min_keyword_length 
                      and not (kw.startswith('["') or kw.startswith('[\'') or kw.startswith('{'))]
            if keywords:
                return keywords[:max_keywords]
    
    return []


def _normalize_multi_prompts_list(prompts: List[str]) -> List[str]:
    """Normalize multi_prompts: strip, drop empty, dedup exact (keep order)."""
    if not prompts:
        return []
    seen = set()
    normalized = []
    for p in prompts:
        if p is None:
            continue
        s = str(p).strip()
        if not s:
            continue
        if s in seen:
            continue
        seen.add(s)
        normalized.append(s)
    return normalized


def _extract_system_collections_from_config(config: Optional[Dict[str, Any]]) -> set:
    """從配置中提取系統儲存區（collection）名稱（模組層輔助函式）"""
    if not config or not isinstance(config, dict):
        return set()
    collections = set()
    api_config = config.get('api', {})
    if 'default_collection' in api_config:
        collections.add(api_config['default_collection'].lower())
    if 'qa_default_collection' in api_config:
        collections.add(api_config['qa_default_collection'].lower())
    if 'add_point_default_collection' in api_config:
        collections.add(api_config['add_point_default_collection'].lower())
    question_analysis = api_config.get('question_analysis', {})
    if 'default_collection' in question_analysis:
        collections.add(question_analysis['default_collection'].lower())
    search_settings = config.get('search_settings', {})
    if 'collection_name' in search_settings:
        collections.add(search_settings['collection_name'].lower())
    crawler = config.get('crawler', {})
    sharepoint = crawler.get('sharepoint', {})
    if 'default_collection' in sharepoint:
        collections.add(sharepoint['default_collection'].lower())
    collections_config = config.get('collections', {})
    auto_create = collections_config.get('auto_create', {})
    for pattern in auto_create.get('patterns', []):
        if isinstance(pattern, str) and '*' not in pattern and '?' not in pattern:
            collections.add(pattern.lower())
    collections.update({'test', 'test_update_flow', 'test_duplicate_path_flow', 'test_batch', 'test_collection_basic', 'test_basic'})
    return collections


def filter_meaningless_tags(tags: List[str], config: Optional[Dict[str, Any]] = None) -> List[str]:
    """
    過濾掉無意義的標籤（模組層函式，可被 txt_parser 等直接呼叫）。
    """
    if not tags:
        return []
    cfg = config if config is not None else m_config
    system_collections = _extract_system_collections_from_config(cfg)
    system_path_keywords = {
        'c:', 'd:', 'e:', 'f:', 'g:', 'h:',
        'program files', 'program files (x86)', 'programdata',
        'users', 'user', 'documents', 'desktop', 'downloads',
        'appdata', 'local', 'roaming', 'temp', 'tmp',
        'windows', 'system32', 'syswow64', 'winnt',
        'usr', 'var', 'etc', 'opt', 'home', 'root', 'bin', 'sbin',
        'lib', 'lib64', 'dev', 'proc', 'sys', 'mnt', 'media',
        'node_modules', '.git', '.vscode', '.idea', '__pycache__',
        'venv', 'env', 'virtualenv', '.venv',
        'system', 'system64', 'drivers',
        'common', 'shared', 'public', 'private',
        'storage', 'log', 'logs', 'cache', 'backup',
        'documentmanager', 'domain', 'test', 'tests', 'test_', 'tmp_',
        '.svn', 'docxparser', 'contextparser', 'source', 'ml_home'
    }
    filtered_tags = []
    for tag in tags:
        if not tag or not isinstance(tag, str):
            continue
        tag_lower = tag.lower().strip()
        if not tag_lower:
            continue
        if tag_lower.isdigit():
            continue
        digit_count = sum(1 for c in tag_lower if c.isdigit())
        if len(tag_lower) > 0 and digit_count / len(tag_lower) > 0.8:
            continue
        if tag_lower in system_path_keywords:
            continue
        if tag_lower.replace(' ', '').replace('_', '').replace('-', '') in system_path_keywords:
            continue
        if len(tag_lower) < 2:
            continue
        if re.match(r'^v?\d+(\.\d+)*$', tag_lower):
            continue
        if tag_lower.startswith('~$') or tag_lower.endswith(('.tmp', '.bak', '.old', '.swp')):
            continue
        if tag_lower in system_collections:
            continue
        filtered_tags.append(tag)
    return filtered_tags


def dedup_multi_prompts_by_llm(
        prompts_by_segment: List[List[str]],
        config: Optional[Dict[str, Any]] = None,
        llm_base_url: Optional[str] = None,
        llm_provider: str = 'remote',
        llm_model: str = 'remote8b',
        logger: Optional[logging.Logger] = None,
        return_meta: bool = False
    ) -> Union[List[List[str]], Tuple[List[List[str]], List[Dict[str, Any]]]]:
    """
        LLM-assisted dedup for multi_prompts (keep by mask indices).
        - First pass: strict dedup (exact match).
        - If LLM fails or returns invalid mask: keep all (after strict dedup).

        IO example (List[List[str]] in, List[List[str]] out):
        input:  [["A","B","B","C"], ["x","y"]]
        output: [["A","B","C"], ["x","y"]]  # after strict + optional LLM mask
    """
    if not prompts_by_segment:
        return []

    cfg = config or m_config
    log = logger or m_logger

    llm_cfg = cfg.get('llm', {}) if isinstance(cfg, dict) else {}
    seg_cfg = llm_cfg.get('segment', {}) if isinstance(llm_cfg, dict) else {}
    dedup_cfg = seg_cfg.get('dedup_multi_prompts', {}) if isinstance(seg_cfg, dict) else {}

    enabled = dedup_cfg.get('enabled', True)
    min_count = int(dedup_cfg.get('min_count', 3))
    batch_size = int(dedup_cfg.get('batch_size', 8))
    max_items = int(dedup_cfg.get('max_items', 40))
    max_keep = int(dedup_cfg.get('max_keep', 0))

    cleaned = [_normalize_multi_prompts_list(p) for p in prompts_by_segment]
    metas = [{'success': False, 'reason': 'skipped'} for _ in cleaned]
    if not enabled:
        for i in range(len(metas)):
            metas[i] = {'success': False, 'reason': 'disabled'}
        return (cleaned, metas) if return_meta else cleaned

    # 讀取 prompt 配置
    prompt_path = cfg.get('prompt_path', None) if isinstance(cfg, dict) else None
    if prompt_path:
        prompt_file = Path(prompt_path) / 'dedup_multi_prompts.json'
    else:
        project_root = Path(__file__).parent.parent.parent
        prompt_file = project_root / 'prompt' / 'dedup_multi_prompts.json'

    try:
        if prompt_file.exists():
            with open(prompt_file, 'r', encoding='utf-8-sig') as f:
                dedup_prompt_cfg = json.load(f)
        else:
            log.warning(f"Prompt 檔案不存在: {prompt_file}")
            return (cleaned, metas) if return_meta else cleaned
    except Exception as e:
        log.warning(f"載入 dedup_multi_prompts prompt 失敗: {e}")
        return (cleaned, metas) if return_meta else cleaned

    system_prompt = dedup_prompt_cfg.get('system_prompt', '')
    user_prompt_template = dedup_prompt_cfg.get('user_prompt_template', '')
    generation_config = dedup_prompt_cfg.get('generation_config', {})

    if not llm_base_url:
        llm_base_url = llm_cfg.get('base_url', 'http://10.1.3.127:7017')

    prompts_list = []
    index_map = []
    for idx, items in enumerate(cleaned):
        if len(items) < min_count:
            metas[idx] = {'success': False, 'reason': 'below_min_count'}
            continue
        if max_items > 0 and len(items) > max_items:
            items = items[:max_items]

        metas[idx] = {'success': False, 'reason': 'queued'}
        items_json = json.dumps(items, ensure_ascii=False)
        user_prompt = user_prompt_template.format(items_json=items_json)
        prompts_list.append(user_prompt)
        index_map.append((idx, items))

    if not prompts_list:
        return (cleaned, metas) if return_meta else cleaned

    def _apply_mask(result_text: str, original_items: List[str]) -> Tuple[List[str], bool, str]:
        if not result_text:
            return original_items, False, 'empty_response'
        match = re.search(r'\{[\s\S]*\}', result_text.strip())
        if not match:
            return original_items, False, 'no_json_object'
        try:
            parsed = json.loads(match.group(0))
        except Exception:
            return original_items, False, 'json_parse_failed'
        if not isinstance(parsed, dict):
            return original_items, False, 'invalid_object'

        mask = parsed.get('mask')
        success = parsed.get('success')
        reason = parsed.get('reason') or ''
        if success is not True:
            return original_items, False, reason or 'success_false'

        if not isinstance(mask, list):
            return original_items, False, 'invalid_mask'
        if len(mask) != len(original_items):
            return original_items, False, 'mask_length_mismatch'
        for m in mask:
            if not isinstance(m, int) or m not in (0, 1):
                return original_items, False, 'invalid_mask_value'
        kept = [x for x, m in zip(original_items, mask) if m == 1]
        if not kept:
            return original_items, False, 'empty_mask'
        return kept, True, reason

    max_tokens = generation_config.get('max_new_tokens', 2000)
    temperature = generation_config.get('temperature', 0.1)

    # 批量請求
    batch_chat_url = f"{llm_base_url.rstrip('/')}/chat/batch"
    results = []
    try:
        for start in range(0, len(prompts_list), max(1, batch_size)):
            batch_prompts = prompts_list[start:start + max(1, batch_size)]
            payload = {
                'prompts': batch_prompts,
                'provider': llm_provider,
                'model': llm_model,
                'max_tokens': max_tokens,
                'temperature': temperature,
                'system_prompt': system_prompt if system_prompt else None,
                'parallel': True,
                'max_batch_size': 190
            }
            response = requests.post(batch_chat_url, json=payload, timeout=300)
            response.raise_for_status()
            batch_result = response.json()
            batch_items = batch_result.get('results', [])

            for item in batch_items:
                if item.get('error'):
                    results.append(None)
                else:
                    result_obj = item.get('result', {})
                    output = result_obj.get('output', '')
                    results.append(output)
    except Exception as e:
        log.warning(f"[dedup_multi_prompts_by_llm] 請求llm 對 multi_prompts 格式去重失敗: {e}")
        for seg_idx, _ in index_map:
            metas[seg_idx] = {'success': False, 'reason': '請求LLM格式去重失敗'}
        return (cleaned, metas) if return_meta else cleaned

    for (seg_idx, original_items), result_text in zip(index_map, results):
        if not result_text:
            metas[seg_idx] = {'success': False, 'reason': 'LLM格式去重回傳空字串'}
            continue
        deduped, ok, reason = _apply_mask(result_text, original_items)
        if ok and reason:
            log.info(f"[dedup_multi_prompts_by_llm] LLM reason: {reason}")
        elif not ok and reason:
            log.info(f"[dedup_multi_prompts_by_llm] LLM reason: {reason}")
        if max_keep > 0:
            deduped = deduped[:max_keep]
        cleaned[seg_idx] = deduped
        metas[seg_idx] = {'success': ok, 'reason': reason}

    for seg_idx, _ in index_map:
        if metas[seg_idx].get('reason') == 'queued':
            metas[seg_idx] = {'success': False, 'reason': 'no_result'}

    return (cleaned, metas) if return_meta else cleaned


def _load_image_explainer_prompt_template(config: Optional[Dict[str, Any]], log: logging.Logger) -> str:
    """載入 image_explainer prompt 模板，失敗時回退內建預設。"""
    default_template = (
        "請根據上下文描述圖片內容，並給 1-2 句總結。"
        "如果上下文中有指示性的描述，是透過圖片來輔助說明的，"
        "那你應該就圖像中的相對位置來補充上下文單純透過文字沒有表達出來的資訊。"
        "而且：如果圖裡面有表格，請以md格式輸出前10行前10欄的內容；"
        "如果圖裡面是座標xy chart，請描述座標軸意義與圖表類型"
    )
    cfg = config or {}
    try:
        project_root = Path(__file__).resolve().parent.parent.parent
        prompt_path = cfg.get('prompt_path', None) if isinstance(cfg, dict) else None
        prompt_file = Path(prompt_path) / 'image_explainer.json' if prompt_path else (project_root / 'prompt' / 'image_explainer.json')

        if not prompt_file.exists():
            log.warning(f"[_analyze_images_via_batch] prompt 檔案不存在，使用內建預設: {prompt_file}")
            return default_template

        with open(prompt_file, 'r', encoding='utf-8-sig') as f:
            raw = f.read().strip()
        if not raw:
            log.warning(f"[_analyze_images_via_batch] prompt 檔案為空，使用內建預設: {prompt_file}")
            return default_template

        loaded = json.loads(raw)
        if isinstance(loaded, str):
            template = loaded.strip()
            return template or default_template
        if isinstance(loaded, dict):
            for key in ('image_prompt_template', 'prompt_template', 'user_prompt_template'):
                val = loaded.get(key)
                if val is not None:
                    template = str(val).strip()
                    if template:
                        return template
        return default_template
    except Exception as e:
        log.warning(f"[_analyze_images_via_batch] 載入 image_explainer prompt 失敗，使用內建預設: {e}")
        return default_template




def analyze_images_via_batch_common(
        segments: List[Dict[str, Any]],
        metadata: Optional[Dict[str, Any]],
        *,
        llm_provider: str = 'openai',
        llm_model: str = 'gpt4o_chat',
        llm_base_url: Optional[str] = None,
        enable_image_llm: bool = True,
        image_context_window: int = 200,
        max_images_per_batch: int = 50,
        image_prompt_template: Optional[str] = None,
        placeholder_pattern: Optional[str] = None,
        placeholder_replacements: Optional[Union[str, List[str]]] = None,
        text_keys: Optional[List[str]] = None,
        config: Optional[Dict[str, Any]] = None,
        logger: Optional[logging.Logger] = None
    ) -> Dict[str, Any]:
    """
    Shared image analysis for parsers. Scans placeholders, calls /chat/batch,
    and inserts summaries back into segments and metadata.
    """
    log = logger or m_logger
    cfg = config or m_config

    if not enable_image_llm:
        log.warning("[_analyze_images_via_batch] 圖像 LLM 分析已停用")
        return {}

    images = (metadata or {}).get("images", []) if isinstance(metadata, dict) else []
    if not images:
        log.debug("[_analyze_images_via_batch] 沒有圖像需要分析")
        return {}

    if not llm_base_url:
        llm_base_url = cfg.get('llm', {}).get('base_url', 'http://10.1.3.127:7017') if isinstance(cfg, dict) else 'http://10.1.3.127:7017'

    if not image_prompt_template:
        image_prompt_template = _load_image_explainer_prompt_template(cfg, log)

    if not placeholder_pattern:
        placeholder_pattern = r'\[IMAGE_PLACEHOLDER_(\d+)\]'
    placeholder_re = re.compile(placeholder_pattern)

    if not text_keys:
        text_keys = ['unit_text', 'call_prompt', 'text']

    if placeholder_replacements is None:
        placeholder_replacements = ["[IMAGE_PLACEHOLDER_{image_id}]"]
    elif isinstance(placeholder_replacements, str):
        placeholder_replacements = [placeholder_replacements]

    def _get_text(segment: Dict[str, Any]) -> str:
        for key in text_keys:
            if key in segment and segment.get(key):
                value = segment.get(key)
                return value if isinstance(value, str) else str(value)
        return ""

    def _build_placeholders(image_id_str: str) -> List[str]:
        placeholders = []
        for template in placeholder_replacements:
            if "{image_id}" in template:
                placeholders.append(template.format(image_id=image_id_str))
            else:
                placeholders.append(template)
        return placeholders

    def _apply_image_analysis_to_segments(image_analysis_map: Dict[str, Any]) -> None:
        if not image_analysis_map:
            return
        for seg_idx, segment in enumerate(segments):
            text_key = None
            for key in text_keys:
                if key in segment and segment.get(key):
                    text_key = key
                    break
            if not text_key:
                continue
            current_text = segment.get(text_key, "")
            current_text = current_text if isinstance(current_text, str) else str(current_text)
            matched_ids = placeholder_re.findall(current_text)
            if not matched_ids:
                continue
            appended_ids: Set[str] = set()
            for match in matched_ids:
                try:
                    iid = int(match)
                except Exception:
                    continue
                iid_key = str(iid)
                analysis = image_analysis_map.get(iid_key)
                if not analysis or not isinstance(analysis, dict):
                    continue
                if "error" in analysis:
                    continue
                summary = analysis.get("summary", "")
                if not summary:
                    continue
                annotation_line = f"[圖像說明] {summary}"
                if annotation_line in current_text:
                    continue
                placeholders = _build_placeholders(iid_key)
                for ph in placeholders:
                    if ph in current_text:
                        current_text = current_text.replace(ph, f"{ph}\n{annotation_line}")
                if iid_key not in appended_ids:
                    segment.setdefault("image_analysis", []).append(analysis)
                    appended_ids.add(iid_key)
            segment[text_key] = current_text

    image_tasks: List[Dict[str, Any]] = []
    seen_image_ids: Set[int] = set()
    placeholder_ref_total = 0

    for seg_idx, segment in enumerate(segments):
        text = _get_text(segment)
        if not text:
            log.warning(f"[_analyze_images_via_batch] 段落 {seg_idx} 無文字內容")
            continue

        matches = placeholder_re.findall(text)
        for match in matches:
            placeholder_ref_total += 1
            try:
                image_id = int(match)
            except Exception:
                continue

            if image_id >= len(images):
                log.warning(f"[_analyze_images_via_batch] 圖像 ID {image_id} 超出範圍（總圖像數: {len(images)}）")
                continue

            if image_id in seen_image_ids:
                continue
            seen_image_ids.add(image_id)

            front_text = ""
            back_text = ""
            if seg_idx > 0:
                prev_text = _get_text(segments[seg_idx - 1])
                front_text = prev_text[-image_context_window:] if len(prev_text) > image_context_window else prev_text
            if seg_idx < len(segments) - 1:
                next_text = _get_text(segments[seg_idx + 1])
                back_text = next_text[:image_context_window] if len(next_text) > image_context_window else next_text

            if front_text or back_text:
                prompt = f"{image_prompt_template}\n上文：{front_text}\n下文：{back_text}"
            else:
                prompt = image_prompt_template

            img_data = images[image_id]
            image_url = None
            if isinstance(img_data, dict):
                if 'base64' in img_data:
                    mime = img_data.get('mime', 'image/jpeg')
                    base64_str = img_data['base64']
                    image_url = f"data:{mime};base64,{base64_str}"
                elif 'image_url' in img_data:
                    image_url = img_data['image_url']

            if not image_url:
                log.warning(f"[_analyze_images_via_batch] 圖像 ID {image_id} 無法取得圖像資料")
                continue

            image_tasks.append({
                'image_id': image_id,
                'image_url': image_url,
                'prompt': prompt,
                'segment_idx': seg_idx,
                'placeholder_match': match
            })

    unique_image_count = len(seen_image_ids)
    existing_ia = metadata.get("image_analysis") if isinstance(metadata, dict) else None
    has_existing_analysis = isinstance(existing_ia, dict) and len(existing_ia) > 0

    if not image_tasks and not has_existing_analysis:
        log.debug("[_analyze_images_via_batch] 沒有找到圖像佔位符")
        return {}

    if image_tasks:
        log.info(
            f"[_analyze_images_via_batch] 佔位符命中={placeholder_ref_total} "
            f"API 任務數={len(image_tasks)} unique_image_count={unique_image_count}"
        )
    elif has_existing_analysis:
        log.info(
            f"[_analyze_images_via_batch] 佔位符命中={placeholder_ref_total} "
            f"沿用既有 image_analysis 條目數={len(existing_ia)} unique_image_count={len(existing_ia)}"
        )

    image_analysis: Dict[str, Any] = {}
    skip_batch_api = bool(has_existing_analysis)

    if skip_batch_api:
        image_analysis = dict(metadata.get("image_analysis") or {})
        log.debug("[_analyze_images_via_batch] 略過批次 API（metadata 已有 image_analysis）")
    else:
        batch_chat_url = f"{llm_base_url.rstrip('/')}/chat/batch"

        for batch_start in range(0, len(image_tasks), max_images_per_batch):
            batch_tasks = image_tasks[batch_start:batch_start + max_images_per_batch]
            payload = {
                'prompts': [task['prompt'] for task in batch_tasks],
                'provider': llm_provider,
                'model': llm_model,
                'max_tokens': 12000,
                'temperature': 0.1,
                'top_p': 0.95,
                'images': [[task['image_url']] for task in batch_tasks]
            }

            try:
                log.debug(f"[_analyze_images_via_batch] 發送批次請求: {len(batch_tasks)} 個圖像")
                resp = requests.post(batch_chat_url, json=payload, timeout=600)
                resp.raise_for_status()
                batch_data = resp.json()

                results = batch_data.get('results', [])
                for i, task in enumerate(batch_tasks):
                    if i < len(results):
                        result_item = results[i]
                        if result_item.get('error'):
                            log.warning(f"[_analyze_images_via_batch] 圖像 ID {task['image_id']} 分析失敗: {result_item['error']}")
                            image_analysis[str(task['image_id'])] = {
                                'error': result_item['error'],
                                'image_id': task['image_id']
                            }
                        else:
                            result = result_item.get('result') or {}
                            output = result.get('output', '')
                            # 圖文解釋：後續在此銜接 has_table / md 表格檢查與重試（見 src.image_explainer_output）
                            output = finalize_image_explainer_text(output)
                            cost = result.get('current_cost_usd', 0)
                            usage = result.get('usage') or {}
                            total_tokens = usage.get('total_tokens', 0) if isinstance(usage, dict) else 0

                            image_analysis[str(task['image_id'])] = {
                                'summary': output,
                                'prompt_used': task['prompt'],
                                'model': llm_model,
                                'cost': cost,
                                'tokens': total_tokens,
                                'image_id': task['image_id'],
                                'segment_idx': task['segment_idx']
                            }
                    else:
                        log.warning(f"[_analyze_images_via_batch] 圖像 ID {task['image_id']} 沒有對應的回覆")
                        image_analysis[str(task['image_id'])] = {
                            'error': 'No response',
                            'image_id': task['image_id']
                        }

                batch_cost = batch_data.get('batch_cost_usd', 0)
                batch_token_summary = batch_data.get('batch_token_summary') or {}
                if not isinstance(batch_token_summary, dict):
                    batch_token_summary = {}
                log.info(f"[_analyze_images_via_batch] 批次 {batch_start//max_images_per_batch + 1} 完成: 費用={batch_cost}, tokens={batch_token_summary.get('total_tokens', 0)}")

            except requests.RequestException as e:
                log.error(f"[_analyze_images_via_batch] 批次請求失敗: {e}")
                for task in batch_tasks:
                    image_analysis[str(task['image_id'])] = {
                        'error': str(e),
                        'image_id': task['image_id']
                    }

        if image_analysis and isinstance(metadata, dict):
            metadata['image_analysis'] = image_analysis

    _apply_image_analysis_to_segments(image_analysis)

    log.info(f"[_analyze_images_via_batch] 完成圖像分析，共 {len(image_analysis)} 個結果")

    return {
        'image_analysis': image_analysis,
        'total_analyzed': len(image_analysis),
        'total_errors': sum(1 for v in image_analysis.values() if 'error' in v)
    }



def segment_by_custom_pattern(text: str, separator: str) -> List[str]:
    """
    根據自定義分隔符分段文本
    
    Args:
        text: 輸入文字
        separator: 分隔符字符串
        
    Returns:
        分段後的文字列表
    """
    if not text:
        return []
    
    if not separator:
        # 如果分隔符為空，返回整個文本作為單一段落
        return [text] if text.strip() else []
    
    # 使用分隔符分割文本
    segments = text.split(separator)
    
    # 過濾掉空段落並去除首尾空白
    result = [seg.strip() for seg in segments if seg.strip()]
    
    return result if result else [text]  # 如果分割後為空，返回原始文本


_MARKDOWN_HEADING_RE = re.compile(r'^(#{1,6})\s+\S')
_BULLET_RE = re.compile(r'^(\s*)([-*+•]+)\s+\S')
_DECIMAL_RE = re.compile(r'^(\d+(?:\.\d+)+)\s+\S')
_DECIMAL_SIMPLE_RE = re.compile(r'^(\d+)[、.)．]\s+\S')
_HYPHEN_RE = re.compile(r'^(\d+(?:-\d+)+)\s+\S')
_ROMAN_RE = re.compile(r'^([IVXLCDM]+)[.)]\s+\S', re.I)
_ALPHA_RE = re.compile(r'^([A-Za-z])[.)]\s+\S')
_CN_NUMERAL_RE = re.compile(r'^([一二三四五六七八九十百千]+)[、．.\)]\s+\S')
_TABLE_SEP_RE = re.compile(r'^\s*\|?[-:\s]+(\|[-:\s]+)+\|?\s*$')
_CODE_FENCE_RE = re.compile(r'^\s*(`{3,}|~{3,})')


def build_code_block_mask(lines: List[str]) -> List[bool]:
    """
    建立 code block 遮罩（``` 或 ~~~ 之間的行皆視為程式碼）
    """
    mask = [False] * len(lines)
    in_block = False
    fence = None
    for i, line in enumerate(lines):
        stripped = line.strip()
        m = _CODE_FENCE_RE.match(stripped)
        if m:
            current_fence = m.group(1)[0]
            mask[i] = True
            if not in_block:
                in_block = True
                fence = current_fence
            else:
                if fence == current_fence:
                    in_block = False
                    fence = None
            continue
        if in_block:
            mask[i] = True
    return mask


def is_table_like_line(line: str) -> bool:
    """
    判斷是否像表格行（Markdown 表格 / 分隔線 / 多欄位）
    """
    s = line.strip()
    if not s:
        return False
    if _TABLE_SEP_RE.match(s):
        return True
    if s.count('|') >= 2:
        cols = [c.strip() for c in s.split('|') if c.strip()]
        return len(cols) >= 2
    return False


def is_code_like_line(line: str) -> bool:
    """
    判斷是否像程式碼或註解行（需避開 Markdown 標題 / 條列）
    """
    s = line.strip()
    if not s:
        return False
    if _MARKDOWN_HEADING_RE.match(s):
        return False
    if _BULLET_RE.match(line):
        return False
    if _DECIMAL_RE.match(s) or _DECIMAL_SIMPLE_RE.match(s) or _HYPHEN_RE.match(s):
        return False
    if _ROMAN_RE.match(s) or _ALPHA_RE.match(s) or _CN_NUMERAL_RE.match(s):
        return False
    if _CODE_FENCE_RE.match(s):
        return True
    if re.match(r'^(//|/\*|\*|#(?!\s))', s):
        return True
    if re.match(r'^\s*#\s*(include|define|ifdef|ifndef|endif)\b', s, re.I):
        return True
    if re.search(r'\b(def|class|return|import|from|const|let|var|public|private|if|for|while|switch|case)\b', s):
        return True
    if re.search(r'[{};]|->|=>', s):
        return True
    return False


def is_list_like_line(line: str) -> bool:
    """
    判斷是否像條列項（用於連續條列密度檢查）
    """
    s = line.strip()
    if not s:
        return False
    return bool(
        _BULLET_RE.match(line)
        or _DECIMAL_SIMPLE_RE.match(s)
        or _DECIMAL_RE.match(s)
        or _HYPHEN_RE.match(s)
        or _ROMAN_RE.match(s)
        or _ALPHA_RE.match(s)
        or _CN_NUMERAL_RE.match(s)
    )


def looks_like_list_run(lines: List[str], idx: int, code_mask: Optional[List[bool]] = None, window: int = 2) -> bool:
    """
    檢查周圍是否為連續條列（避免把條列當成標題）
    """
    count = 0
    start = max(0, idx - window)
    end = min(len(lines), idx + window + 1)
    for j in range(start, end):
        if j == idx:
            continue
        if code_mask and code_mask[j]:
            continue
        line = lines[j]
        s = line.strip()
        if not s:
            continue
        if is_table_like_line(s) or is_code_like_line(line):
            continue
        if is_list_like_line(line):
            count += 1
    return count >= 2


def build_marker_line_index(matches: List[Dict[str, Any]]) -> Dict[str, Dict[str, List[int]]]:
    """
    建立 marker -> line_no 的索引，方便序號一致性檢查
    """
    idx: Dict[str, Dict[str, List[int]]] = {}
    for m in matches:
        m_type = m.get('type')
        marker = str(m.get('marker') or '')
        line_no = int(m.get('line_no') or 0)
        if not m_type or not marker or line_no <= 0:
            continue
        idx.setdefault(m_type, {}).setdefault(marker, []).append(line_no)
    return idx


def _merge_heading_tags_by_orders(
        orders: List[int],
        heading_tags_by_order: Optional[Dict[int, List[str]]],
        max_tag_len: int = 50
    ) -> List[str]:
    """
    將同一段落內多個 order 的 tags 合併（保留順序、去重）
    """
    if not orders or not heading_tags_by_order:
        return []
    merged: List[str] = []
    for order in orders:
        for tag in heading_tags_by_order.get(order, []) or []:
            if tag is None:
                continue
            if not isinstance(tag, str):
                tag = str(tag)
            if max_tag_len is not None and len(tag) > max_tag_len:
                continue
            if tag and tag not in merged:
                merged.append(tag)
    return merged


def _strip_heading_marker(text: str) -> str:
    """
    移除常見標題符號，只保留標題內容文字
    """
    if not text:
        return ''
    raw = text.rstrip('\r\n')
    s = raw.strip()
    if not s:
        return ''

    # 方括號標題：只能在欄位 0
    if raw.startswith('['):
        m = re.match(r'^\[(\d+(?:[.-]\d+)*)\]\s*(.*)$', raw)
        if m:
            return (m.group(2) or '').strip()
    if raw.startswith('【'):
        m = re.match(r'^【(\d+(?:[.-]\d+)*)】\s*(.*)$', raw)
        if m:
            return (m.group(2) or '').strip()

    # Markdown 標題
    m = re.match(r'^(#{1,6})\s+(.+)$', s)
    if m:
        return m.group(2).strip()

    # 中文章節
    m = re.match(r'^第([一二三四五六七八九十百千0-9]+)([章節條款項])\s*(.*)$', s)
    if m:
        return (m.group(3) or '').strip()

    # 中文數字序號
    m = re.match(r'^([一二三四五六七八九十百千]+)[、．.\)]\s*(.*)$', s)
    if m:
        return (m.group(2) or '').strip()

    # 多級數字/簡單數字/破折號
    m = re.match(r'^(\d+(?:\.\d+)+)\s*(.*)$', s)
    if m:
        return (m.group(2) or '').strip()
    m = re.match(r'^(\d+)[、.)．]\s*(.*)$', s)
    if m:
        return (m.group(2) or '').strip()
    m = re.match(r'^(\d+(?:-\d+)+)\s*(.*)$', s)
    if m:
        return (m.group(2) or '').strip()

    # 羅馬數字/英文字母
    m = re.match(r'^([IVXLCDM]+)[.)]\s*(.*)$', s, re.I)
    if m:
        return (m.group(2) or '').strip()
    m = re.match(r'^([A-Za-z])[.)]\s*(.*)$', s)
    if m:
        return (m.group(2) or '').strip()

    return s


def _build_heading_tags_by_order(unit_paras: List[Dict[str, Any]]) -> Dict[int, List[str]]:
    """
    依據 unit_paras 的 structure_chars 建立 order -> tags 的對應表
    """
    heading_tags_by_order: Dict[int, List[str]] = {}
    if not unit_paras:
        return heading_tags_by_order

    heading_stack: List[str] = []
    for para in sorted(unit_paras, key=lambda p: p.get('order', 0)):
        structure_chars = para.get('structure_chars', []) or []
        heading_info = next((sc for sc in structure_chars if sc.get('type') == 'heading'), None)
        if heading_info:
            level = heading_info.get('level', 1)
            try:
                level = int(level)
            except (ValueError, TypeError):
                level = 1
            if level <= 0:
                level = 1
            heading_text = _strip_heading_marker(para.get('unit_text', '') or '')
            if heading_text:
                if len(heading_stack) < level:
                    heading_stack.extend([None] * (level - len(heading_stack)))
                heading_stack[level - 1] = heading_text
                for idx in range(level, len(heading_stack)):
                    heading_stack[idx] = None

        heading_tags_by_order[para.get('order', 0)] = [t for t in heading_stack if t]

    return heading_tags_by_order


def _marker_prefixes(marker: str, sep: str) -> List[str]:
    parts = marker.split(sep)
    prefixes = []
    for i in range(1, len(parts)):
        prefixes.append(sep.join(parts[:i]))
    return prefixes


def _has_prior_marker(marker_index: Dict[str, Dict[str, List[int]]], types: List[str], marker: str, line_no: int, window: int) -> bool:
    for t in types:
        lines = marker_index.get(t, {}).get(marker, [])
        for ln in lines:
            if ln < line_no and (window is None or (line_no - ln) <= window):
                return True
    return False


def _has_prior_cn_chapter(marker_index: Dict[str, Dict[str, List[int]]], digits: str, line_no: int, window: int) -> bool:
    if not digits:
        return False
    for marker, lines in marker_index.get('cn_chapter', {}).items():
        if digits in marker:
            for ln in lines:
                if ln < line_no and (window is None or (line_no - ln) <= window):
                    return True
    return False


def pass_serial_consistency(match: Dict[str, Any], marker_index: Dict[str, Dict[str, List[int]]], window: int = 5) -> bool:
    """
    序號一致性檢查：避免孤立的 1.1 / 1-1-1 被誤判為標題
    """
    m_type = match.get('type')
    marker = str(match.get('marker') or '')
    line_no = int(match.get('line_no') or 0)
    if m_type not in ('decimal', 'hyphen'):
        return True
    sep = '.' if m_type == 'decimal' else '-'
    if sep not in marker:
        return True
    prefixes = _marker_prefixes(marker, sep)
    for prefix in reversed(prefixes):
        if _has_prior_marker(marker_index, ['decimal', 'decimal_simple', 'hyphen'], prefix, line_no, window):
            return True
        digits = ''.join(re.findall(r'\d+', prefix))
        if _has_prior_cn_chapter(marker_index, digits, line_no, window):
            return True
    return False


# FP 過濾：版本號、內文引用、問卷選項、產品型號、清單、枚舉 等關鍵字
_FP_NEGATIVE_PHRASES = (
    '版本號', '節點代碼', '不是章節標題', '非章節標題',
    '為步驟概述', '問卷選項', '不是標題', '產品型號',
    '（只是清單）', '只是枚舉', '只是清單',
)
_FP_CN_CHAPTER_REF_RE = re.compile(r'^第[一二三四五六七八九十百千0-9]+[章節條款項]為')  # 第1條為、第2章為
_FP_HYPHEN_CODE_RE = re.compile(r'\d+(?:-\d+)+-[A-Za-z0-9]')  # 1-1-1-ABC
_FP_MULTI_ALPHA_RE = re.compile(r'^[A-Za-z]\.\s+[A-Za-z]\.\s+[A-Za-z]\.')  # A. B. C.
_FP_MULTI_CN_NUM_RE = re.compile(r'^[一二三四五六七八九十]、[一二三四五六七八九十]、[一二三四五六七八九十]')  # 一、二、三


def _is_likely_fp_by_context(match: Dict[str, Any], lines: List[str], line_text: str) -> bool:
    """
    依上下文與內容關鍵字，過濾易誤判為標題的 FP（版本號、內文引用、問卷選項等）。
    取代 LLM 驗證，純邏輯判斷。
    """
    m_type = (match.get('type') or '').lower()
    text = (line_text or match.get('text') or match.get('line') or '').strip()

    # 上下文：前 1～3 行含「應視為非標題」等 → 該區塊為負例
    line_no = int(match.get('line_no') or 0)
    for j in range(max(1, line_no - 3), line_no):
        idx = j - 1  # line_no 為 1-based
        if 0 <= idx < len(lines):
            prev = (lines[idx] or '').strip()
            if '應視為非標題' in prev or '誤判風險' in prev:
                return True

    # 內容關鍵字
    for phrase in _FP_NEGATIVE_PHRASES:
        if phrase in text:
            return True

    # 內文引用：第X條為、第X章為
    if m_type == 'cn_chapter' and _FP_CN_CHAPTER_REF_RE.search(text):
        return True

    # 產品型號：1-1-1-ABC
    if m_type == 'hyphen' and _FP_HYPHEN_CODE_RE.search(text):
        return True

    # 枚舉：A. B. C. 這行只是枚舉
    if m_type == 'alpha' and (_FP_MULTI_ALPHA_RE.search(text) or '只是枚舉' in text):
        return True

    # 一、二、三 為步驟概述
    if m_type == 'cn_numeral' and _FP_MULTI_CN_NUM_RE.search(text):
        return True

    return False


def filter_heading_fp_by_context(
        lines: List[str],
        matches: List[Dict[str, Any]],
    ) -> List[Dict[str, Any]]:
    """
    以邏輯規則過濾 FP（版本號、內文引用、問卷選項、產品型號、清單、枚舉等）。
    取代 LLM 驗證。
    """
    if not matches:
        return matches
    filtered = []
    for m in matches:
        line_no = int(m.get('line_no') or 0)
        line_text = (lines[line_no - 1] if 1 <= line_no <= len(lines) else '') or m.get('line', '') or m.get('text', '')
        if _is_likely_fp_by_context(m, lines, line_text):
            continue
        filtered.append(m)
    return filtered


def _load_verify_heading_config(config: Optional[Dict[str, Any]] = None) -> Optional[Dict[str, Any]]:
    """載入標題語境驗證 prompt 配置"""
    try:
        base = Path(__file__).parent.parent.parent
        prompt_path = Path(config.get('prompt_path', base / 'prompt')) if config else base / 'prompt'
        cfg_file = prompt_path / 'verify_heading_llm.json'
        if cfg_file.exists():
            with open(cfg_file, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return None


def _parse_verify_heading_result(response: Optional[str]) -> bool:
    """解析 LLM 回傳 JSON：is_heading=1=True 保留，0/空/錯誤=False 過濾"""
    if not response or not isinstance(response, str):
        return False
    s = response.strip()
    try:
        # 嘗試直接解析 JSON
        if s.startswith('{'):
            obj = json.loads(s)
            if isinstance(obj, dict):
                v = obj.get('is_heading')
                if v in (1, True, '1'):
                    return True
                return False
        # 嘗試從 markdown 代碼塊提取
        m = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', s, re.DOTALL)
        if m:
            obj = json.loads(m.group(1))
            if isinstance(obj, dict):
                v = obj.get('is_heading')
                if v in (1, True, '1'):
                    return True
                return False
        # 嘗試提取第一個 JSON 物件
        start = s.find('{')
        end = s.rfind('}')
        if start != -1 and end > start:
            obj = json.loads(s[start:end + 1])
            if isinstance(obj, dict):
                v = obj.get('is_heading')
                if v in (1, True, '1'):
                    return True
                return False
    except (json.JSONDecodeError, ValueError, TypeError):
        pass
    return False

# TODO: 產生太多FN，準備棄用
def verify_heading_matches_by_llm(
        lines: List[str],
        matches: List[Dict[str, Any]],
        parallel_executor: Any,
        config: Dict[str, Any],
        llm_config: Optional[Dict[str, Any]] = None,
    ) -> List[Dict[str, Any]]:
    """
    將規則偵測到的標題候選送 batch_chat，過濾語境上非標題者。
    """
    if not matches or not parallel_executor or not config:
        return matches

    verify_cfg = llm_config or _load_verify_heading_config(config)
    if not verify_cfg:
        return matches

    user_tpl = verify_cfg.get('user_prompt_template', '')
    system_prompt = verify_cfg.get('system_prompt', '')
    gen_cfg = verify_cfg.get('generation_config', {})
    max_tokens = int(gen_cfg.get('max_new_tokens', 10))
    temperature = float(gen_cfg.get('temperature', 0.1))

    llm_cfg = config.get('llm', {})
    provider = llm_cfg.get('chat_provider', 'remote')
    model = llm_cfg.get('chat_model', 'remote8b')

    prompts_list = []
    for m in matches:
        line_no = int(m.get('line_no') or 0)
        line_text = (lines[line_no - 1] if 1 <= line_no <= len(lines) else m.get('line', '')) or m.get('text', '')
        if not user_tpl:
            continue
        # 用 replace 避免 line_text 含 {is_heading} 等被 format 解析導致 KeyError
        prompt = user_tpl.replace('{line}', line_text[:500])
        prompts_list.append({
            'prompt': prompt,
            'system_prompt': system_prompt,
            'provider': provider,
            'model': model,
            'max_tokens': max_tokens,
            'temperature': temperature,
        })

    if not prompts_list:
        return matches

    try:
        results = parallel_executor.batch_chat(
            prompts_list,
            provider=provider,
            model=model,
            system_prompt=system_prompt,
            max_tokens=max_tokens,
            temperature=temperature,
        )
    except Exception as e:
        m_logger.warning(f"標題語境驗證 batch_chat 失敗: {e}，保留全部候選")
        return matches

    verified = []
    for m, resp in zip(matches, results or []):
        if _parse_verify_heading_result(resp):
            verified.append(m)
    return verified


def apply_text_heading_fallback(
        unit_paras: List[Dict[str, Any]],
        llm_service: Optional[Any] = None,
        config: Optional[Dict[str, Any]] = None,
    ) -> int:
    """
    文字規則備援：將偵測到的標題結果補到 unit_paras.structure_chars。
    若提供 llm_service 與 config， detect_composite_heading_structure 內部會以 batch_chat 驗證語境。

    Args:
        unit_paras: extract() 的 unit_paras 結構
        llm_service: LLM 服務（可選，供 detect_composite_heading_structure 做語境驗證）
        config: 系統配置（可選）

    Returns:
        int: 新增的 heading 數量
    """
    if not unit_paras:
        return 0

    # 將每個 unit_text 壓成單行，避免換行破壞行號對應
    line_texts = []
    for unit in unit_paras:
        raw_text = unit.get('unit_text', '') or ''
        line_texts.append(raw_text.replace('\r', ' ').replace('\n', ' '))

    det_text = '\n'.join(line_texts)
    det = detect_composite_heading_structure(det_text, llm_service=llm_service, config=config)
    matches = det.get('matches', [])

    det_map: Dict[int, Dict[str, Any]] = {}
    for m in matches:
        try:
            line_no = int(m.get('line_no') or 0)
        except Exception:
            continue
        if line_no <= 0 or line_no > len(unit_paras):
            continue
        prev = det_map.get(line_no)
        if not prev or int(m.get('level') or 0) > int(prev.get('level') or 0):
            det_map[line_no] = m

    added_heading_count = 0
    for idx, unit in enumerate(unit_paras, start=1):
        structure_chars = unit.get('structure_chars', []) or []
        # 已有標題就不再補判
        if any(sc.get('type') == 'heading' for sc in structure_chars):
            continue
        # 表格/圖片不補判
        if any(sc.get('type') in ('table', 'image') for sc in structure_chars):
            continue
        m = det_map.get(idx)
        if not m:
            continue
        structure_chars.append({
            'type': 'heading',
            'level': int(m.get('level') or 0),
            'style': 'text_rule',
            'outline_level': -1
        })
        unit['structure_chars'] = structure_chars
        added_heading_count += 1

    return added_heading_count


def detect_composite_heading_structure(
        text: str,
        max_lines: int = 2000,
        llm_service: Optional[Any] = None,
        config: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
    """
    快速偵測多種標題/條列結構（例如：一、 1. 1-1 1.1 * ** - -- # ##）。
    若提供 llm_service 與 config，會再以 batch_chat 驗證語境是否為標題。
    回傳偵測到的樣式統計與逐行命中結果。
    """
    if not text:
        m_logger.warning("[detect_composite_heading_structure] 輸入文字為空，無法偵測標題規律")
        return {'patterns': [], 'matches': []}

    lines = text.splitlines()
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines]

    m_logger.debug(f"[detect_composite_heading_structure] 行數={len(lines)}, max_lines={max_lines}")

    code_mask = build_code_block_mask(lines)

    setext_equals_pattern = re.compile(r'^={3,}$')
    setext_dashes_pattern = re.compile(r'^-{3,}$')

    patterns = [
        ('markdown', re.compile(r'^(#{1,6})\s+(.+)$'),
         lambda m, s: len(m.group(1)),
         lambda m, s: m.group(1),
         lambda m, s: m.group(2).strip()),
        ('cn_chapter', re.compile(r'^第([一二三四五六七八九十百千0-9]+)([章節條款項])\s*(.*)$'),
         lambda m, s: {'章': 1, '節': 2, '條': 3, '款': 4, '項': 5}.get(m.group(2), 1),
         lambda m, s: f"第{m.group(1)}{m.group(2)}",
         lambda m, s: m.group(3).strip()),
        ('cn_numeral', re.compile(r'^([一二三四五六七八九十百千]+)[、．.\)]\s*(.*)$'),
         lambda m, s: 1,
         lambda m, s: m.group(1),
         lambda m, s: m.group(2).strip()),
        ('decimal', re.compile(r'^(\d+(?:\.\d+)+)\s*(.*)$'),
         lambda m, s: m.group(1).count('.') + 1,
         lambda m, s: m.group(1),
         lambda m, s: m.group(2).strip()),
        ('decimal_simple', re.compile(r'^(\d+)[、.)．]\s*(.*)$'),
         lambda m, s: 1,
         lambda m, s: m.group(1),
         lambda m, s: m.group(2).strip()),
        ('bracket_num', re.compile(r'^\[(\d+(?:[.-]\d+)*)\]\s*(.*)$'),
         lambda m, s: 1,
         lambda m, s: f"[{m.group(1)}]",
         lambda m, s: m.group(2).strip()),
        ('bracket_num_fw', re.compile(r'^【(\d+(?:[.-]\d+)*)】\s*(.*)$'),
         lambda m, s: 1,
         lambda m, s: f"【{m.group(1)}】",
         lambda m, s: m.group(2).strip()),
        ('hyphen', re.compile(r'^(\d+(?:-\d+)+)\s*(.*)$'),
         lambda m, s: m.group(1).count('-') + 1,
         lambda m, s: m.group(1),
         lambda m, s: m.group(2).strip()),
        ('roman', re.compile(r'^([IVXLCDM]+)[.)]\s*(.*)$', re.I),
         lambda m, s: 1,
         lambda m, s: m.group(1),
         lambda m, s: m.group(2).strip()),
        ('alpha', re.compile(r'^([A-Za-z])[.)]\s*(.*)$'),
         lambda m, s: 1,
         lambda m, s: m.group(1),
         lambda m, s: m.group(2).strip()),
        # bullet 類型不偵測為標題（僅視為條列項目）
    ]

    raw_matches: List[Dict[str, Any]] = []
    prev_line = ''
    prev_line_no = None

    for idx, line in enumerate(lines, start=1):
        stripped = line.strip()
        if not stripped:
            prev_line = ''
            prev_line_no = None
            continue

        if setext_equals_pattern.match(stripped) or setext_dashes_pattern.match(stripped):
            if prev_line:
                level = 1 if setext_equals_pattern.match(stripped) else 2
                raw_matches.append({
                    'line_no': prev_line_no,
                    'line': prev_line,
                    'type': 'setext',
                    'level': level,
                    'marker': '===' if level == 1 else '---',
                    'text': prev_line
                })
            prev_line = ''
            prev_line_no = None
            continue

        found = False
        for name, regex, level_fn, marker_fn, text_fn in patterns:
            match_line = stripped
            if name in ('bracket_num', 'bracket_num_fw'):
                if line != line.lstrip():
                    continue
                match_line = line
            m = regex.match(match_line)
            if not m:
                continue
            level = level_fn(m, match_line)
            marker = marker_fn(m, match_line)
            text_part = text_fn(m, match_line)
            raw_matches.append({
                'line_no': idx,
                'line': stripped,
                'type': name,
                'level': level,
                'marker': marker,
                'text': text_part
            })
            found = True
            break

        prev_line = stripped if not found else ''
        prev_line_no = idx if not found else None

    marker_index = build_marker_line_index(raw_matches)
    matches: List[Dict[str, Any]] = []
    stats: Dict[str, int] = {}
    filter_stats = {
        'empty': 0,
        'code_mask': 0,
        'table_like': 0,
        'code_like': 0
    }

    for m in raw_matches:
        line_no = int(m.get('line_no') or 0)
        if line_no <= 0 or line_no > len(lines):
            continue
        line = lines[line_no - 1]
        s = line.strip()
        if not s:
            filter_stats['empty'] += 1
            continue
        # 程式碼區塊內的行（含 # 註解）不當成 markdown 標題
        if code_mask[line_no - 1]:
            filter_stats['code_mask'] += 1
            continue
        if is_table_like_line(s):
            filter_stats['table_like'] += 1
            continue
        if is_code_like_line(line):
            filter_stats['code_like'] += 1
            continue

        m_type = m.get('type')

        matches.append(m)
        stats[m_type] = stats.get(m_type, 0) + 1

    # FP 過濾：以邏輯規則取代 LLM，過濾版本號、內文引用、問卷選項、產品型號、清單、枚舉等
    if matches:
        before_cnt = len(matches)
        matches = filter_heading_fp_by_context(lines, matches)
        after_cnt = len(matches)
        stats = {}
        for m in matches:
            t = m.get('type')
            stats[t] = stats.get(t, 0) + 1
        if before_cnt != after_cnt:
            m_logger.info(f"[detect_composite_heading_structure] 上下文過濾：{before_cnt} -> {after_cnt}")

    patterns_summary = [{'type': k, 'count': v} for k, v in sorted(stats.items(), key=lambda x: (-x[1], x[0]))]
    if not raw_matches:
        sample_lines = [ln.strip() for ln in lines if ln.strip()][:5]
        m_logger.info(f"[detect_composite_heading_structure] 未偵測到任何符合規則的行，樣本={sample_lines}")
    elif raw_matches and not matches:
        m_logger.info(f"[detect_composite_heading_structure] 原始命中 {len(raw_matches)} 筆，但全部被過濾")
        m_logger.info(f"[detect_composite_heading_structure] 過濾統計: {filter_stats}")
    else:
        m_logger.debug(f"[detect_composite_heading_structure] 命中統計: {patterns_summary}")
    return {
        'patterns': patterns_summary,
        'matches': matches
    }


def segment_by_paragraphs(text: str) -> List[str]:
    """
    按段落分段（使用雙換行符 \n\n 作為分隔符）
    
    Args:
        text: 輸入文字
        
    Returns:
        分段後的文字列表
    """
    if not text:
        return []
    
    paragraphs = text.split('\n\n')
    result = [para.strip() for para in paragraphs if para.strip()]
    return result if result else [text]

def segment_by_lines(text: str) -> List[str]:
    """
    按行分段
    
    Args:
        text: 輸入文字
        
    Returns:
        分段後的文字列表
    """
    if not text:
        return []
    
    lines = text.split('\n')
    result = [line.strip() for line in lines if line.strip()]
    return result if result else [text]

def clean_llm_keyword_response(response: str) -> str:
    """
    清洗 LLM 關鍵詞提取回應，移除特殊標記和描述性文字
    
    此函數用於處理 LLM 回應中常見的問題：
    1. 移除 LLM 特殊標記（如 <|eot_id|>）
    2. 移除描述性文字前綴（如 "以下是從該段落中提取的 5 個關鍵詞或關鍵片語："）
    3. 正規化 JSON 格式（處理多餘空格）
    
    Args:
        response: LLM 原始回應文本
        
    Returns:
        清洗後的文本
    """
    if not response or not isinstance(response, str):
        return ""
    
    response_clean = response.strip()
    
    # 1. 移除 LLM 特殊標記
    special_tokens = ['<|eot_id|>', '<|end_of_text|>', '<|endoftext|>', '<|im_end|>', '<|end|>']
    for token in special_tokens:
        response_clean = response_clean.replace(token, '')
    
    response_clean = response_clean.strip()
    
    # 2. 移除描述性文字前綴
    # 策略：如果找到 JSON 陣列的開始符號 [，檢查 [ 之前的文字是否為描述性文字
    json_start_idx = response_clean.find('[')
    if json_start_idx > 0:
        prefix = response_clean[:json_start_idx].strip()
        # 檢查前綴是否包含常見的說明性詞語
        explanation_keywords = ['以下', '關鍵詞', '關鍵片語', '提取', '如下', '包括', '代表', '主旨', '列表']
        if any(kw in prefix for kw in explanation_keywords):
            response_clean = response_clean[json_start_idx:]
    
    # 3. 使用正則表達式移除描述性文字模式（更全面的匹配）
    descriptive_patterns = [
        r'以下是從.*?提取.*?關鍵詞.*?：?\s*',
        r'從.*?中提取.*?關鍵詞.*?：?\s*',
        r'關鍵詞.*?如下.*?：?\s*',
        r'提取.*?關鍵詞.*?：?\s*',
        r'以下是.*?關鍵詞.*?：?\s*',
        r'關鍵詞.*?包括.*?：?\s*'
    ]
    
    for pattern in descriptive_patterns:
        response_clean = re.sub(pattern, '', response_clean, flags=re.IGNORECASE)
    
    response_clean = response_clean.strip()
    
    # 4. 正規化 JSON 格式（處理多餘空格和轉義字符）
    # 將常見的全形標點轉換為半形
    response_clean = response_clean.replace('，', ',').replace('：', ':').replace('；', ';')
    
    # 查找 JSON 陣列的開始和結束位置
    json_start_idx = response_clean.find('[')
    json_end_idx = response_clean.rfind(']')
    
    # 檢查是否有不完整的 JSON（有 [ 但沒有對應的 ]）
    if json_start_idx != -1:
        # 嘗試找到配對的 ]
        bracket_count = 0
        actual_end_idx = -1
        for i in range(json_start_idx, len(response_clean)):
            if response_clean[i] == '[':
                bracket_count += 1
            elif response_clean[i] == ']':
                bracket_count -= 1
                if bracket_count == 0:
                    actual_end_idx = i
                    break
        
        # 如果找到配對的 ]，使用它；否則使用最後一個 ]
        if actual_end_idx != -1:
            json_end_idx = actual_end_idx
        elif json_end_idx == -1 or json_end_idx < json_start_idx:
            # 沒有找到配對的 ]，這是不完整的 JSON
            # 嘗試從不完整的 JSON 中提取關鍵詞
            json_part = response_clean[json_start_idx:]
            # 移除開頭的 `[` 和可能的 `"`
            json_part = json_part.lstrip('[').strip()
            # 移除轉義引號
            json_part = json_part.replace('\\"', '"')
            # 提取所有引號內的內容
            quoted_keywords = re.findall(r'"([^"]+)"', json_part)
            if quoted_keywords:
                # 如果找到引號內的關鍵詞，返回一個有效的 JSON 陣列
                cleaned_keywords = [kw.strip() for kw in quoted_keywords if kw.strip()]
                cleaned_keywords = DFP.m_uniq_thru_set(cleaned_keywords)
                if cleaned_keywords:
                    return json.dumps(cleaned_keywords, ensure_ascii=False)
            # 如果沒有找到引號內的內容，嘗試提取第一個引號後的內容
            first_quote_match = re.search(r'^["\']?([^"\',\]]+)', json_part)
            if first_quote_match:
                potential_keyword = first_quote_match.group(1).strip()
                if potential_keyword:
                    return json.dumps([potential_keyword], ensure_ascii=False)
            # 如果都失敗，返回空字符串
            return ""
    
    if json_start_idx != -1 and json_end_idx != -1 and json_end_idx > json_start_idx:
        # 提取 JSON 陣列部分（只取第一個完整的 JSON 陣列，忽略後面的內容）
        json_part = response_clean[json_start_idx:json_end_idx + 1]
        
        # 修復 JSON 陣列內的格式問題：
        # 1. 先將所有 `\\"` 轉換為 `"`（這是關鍵步驟，必須先執行）
        json_part = json_part.replace('\\"', '"')
        
        # 2. 移除所有換行符號和製表符
        json_part = json_part.replace('\n', '').replace('\r', '').replace('\t', ' ')
        
        # 3. 將中文頓號 `、` 轉換為逗號 `,`（在引號之間）
        # 使用正則表達式確保只在引號之間替換頓號
        json_part = re.sub(r'"\s*、\s*"', '","', json_part)
        # 如果還有剩餘的頓號（不在引號之間），也轉換為逗號
        json_part = json_part.replace('、', ',')
        
        # 4. 清理 JSON 陣列中逗號周圍的額外空格
        # 處理標準格式：`" , "` -> `","`
        json_part = re.sub(r'"\s*,\s*"', '","', json_part)
        
        # 5. 清理多餘的空格（但保留引號內的空格）
        json_part = re.sub(r'\s+', ' ', json_part)  # 將多個空格合併為一個
        json_part = re.sub(r'\s*\[\s*', '[', json_part)  # 清理 [ 周圍的空格
        json_part = re.sub(r'\s*\]\s*', ']', json_part)  # 清理 ] 周圍的空格
        
        # 6. 處理包含描述性文字的元素（例如：'以下是從段落中提取出的 5 個關鍵詞或關鍵片語：\n\n["資訊需求"'）
        # 如果第一個元素包含描述性文字和嵌套的 JSON 陣列，嘗試提取嵌套的陣列
        # 匹配模式：描述性文字 + \n\n + [開頭的 JSON 陣列
        nested_array_match = re.search(r'\["([^"]+)"', json_part)
        if nested_array_match:
            # 找到嵌套的陣列開始位置
            nested_start = json_part.find('["')
            if nested_start > 0:
                # 檢查前面是否有描述性文字
                prefix_before_nested = json_part[:nested_start].strip()
                explanation_keywords = ['以下', '關鍵詞', '關鍵片語', '提取', '如下', '包括', '代表', '主旨', '列表']
                if any(kw in prefix_before_nested for kw in explanation_keywords):
                    # 提取從嵌套陣列開始到結尾的部分
                    json_part = json_part[nested_start:]
        
        # 只返回 JSON 陣列部分，忽略後面的所有內容
        response_clean = json_part
    else:
        # 如果沒有找到有效的 JSON 陣列，清理所有轉義字符
        # 移除開頭的 `[\` 和結尾的 `"]`
        response_clean = re.sub(r'^\[\s*\\', '[', response_clean)
        response_clean = re.sub(r'\\\s*\]$', ']', response_clean)
        # 將 `\"` 轉換為 `"`（如果不在有效的 JSON 結構內）
        response_clean = response_clean.replace('\\"', '"')
    
    return response_clean.strip()


def get_storage_roots():
    path_sep = os.sep
    # 準備系統暫存/儲存根路徑，避免將這些路徑寫入 os_tags
    storage_roots: List[str] = []
    try:
        storage_config = m_config.get('storage', {}) if isinstance(m_config, dict) else {}
    except Exception:
        storage_config = {}

    for key in ('temp_path', 'base_path'):
        storage_path = storage_config.get(key)
        if not storage_path:
            continue
        try:
            normalized_root = os.path.normcase(os.path.abspath(storage_path)).rstrip(path_sep)
            storage_roots.append(normalized_root)
        except Exception:
            continue
    return storage_roots
m_storage_roots = get_storage_roots()


class EncodingHandler:
    """
    編碼處理類別
    
    提供多種編碼的解碼方法，用於處理不同編碼格式的文件內容。
    特別針對中文編碼（gb2312, gbk, gb18030）提供錯誤處理策略。
    """
    
    def __init__(self, logger=None):
        """
        初始化編碼處理器
        
        Args:
            logger: 日誌記錄器（可選）
        """
        self.logger = logger or m_logger
        # 定義編碼優先順序（從最寬鬆到最嚴格）
        self.chinese_encodings = ['big5', 'gbk', 'gb18030', 'gb2312', 'utf-8']
        self.common_encodings = ['utf-8', 'big5', 'gbk', 'latin-1', 'cp1252']
        # 支援的編碼列表，包含越南文編碼
        self.supported_encodings = [
            'utf-8',
            'utf-8-sig',  # UTF-8 with BOM
            'windows-1258',  # 越南文
            'big5',  # 繁體中文
            'gb2312',  # 簡體中文
            'gbk',  # 簡體中文擴展
            'gb18030',  # 簡體中文完整
            'windows-1252',  # 西歐語言
            'iso-8859-1',  # Latin-1
            'cp1252',  # Windows-1252 別名
            'latin1',
        ]
        
    def decode_gb2312(self, data: Union[bytes, str], errors: str = 'ignore') -> str:
        """
        解碼 gb2312 編碼的數據
        
        Args:
            data: 要解碼的數據（bytes 或 str）
            errors: 錯誤處理策略（'ignore', 'replace', 'strict'）
            
        Returns:
            解碼後的字符串
        """
        if isinstance(data, str):
            return data
        
        try:
            return data.decode('gb2312', errors=errors)
        except (UnicodeDecodeError, UnicodeError) as e:
            self.logger.warning(f"gb2312 解碼失敗: {e}，嘗試使用 gbk")
            return self.decode_gbk(data, errors)
    
    def decode_gbk(self, data: Union[bytes, str], errors: str = 'ignore') -> str:
        """
        解碼 gbk 編碼的數據
        
        Args:
            data: 要解碼的數據（bytes 或 str）
            errors: 錯誤處理策略（'ignore', 'replace', 'strict'）
            
        Returns:
            解碼後的字符串
        """
        if isinstance(data, str):
            return data
        
        try:
            return data.decode('gbk', errors=errors)
        except (UnicodeDecodeError, UnicodeError) as e:
            self.logger.warning(f"gbk 解碼失敗: {e}，嘗試使用 gb18030")
            return self.decode_gb18030(data, errors)
    
    def decode_gb18030(self, data: Union[bytes, str], errors: str = 'ignore') -> str:
        """
        解碼 gb18030 編碼的數據
        
        Args:
            data: 要解碼的數據（bytes 或 str）
            errors: 錯誤處理策略（'ignore', 'replace', 'strict'）
            
        Returns:
            解碼後的字符串
        """
        if isinstance(data, str):
            return data
        
        try:
            return data.decode('gb18030', errors=errors)
        except (UnicodeDecodeError, UnicodeError) as e:
            self.logger.warning(f"gb18030 解碼失敗: {e}，嘗試使用 utf-8")
            return self.decode_utf8(data, errors)
    
    def decode_utf8(self, data: Union[bytes, str], errors: str = 'ignore') -> str:
        """
        解碼 utf-8 編碼的數據
        
        Args:
            data: 要解碼的數據（bytes 或 str）
            errors: 錯誤處理策略（'ignore', 'replace', 'strict'）
            
        Returns:
            解碼後的字符串
        """
        if isinstance(data, str):
            return data
        
        try:
            return data.decode('utf-8', errors=errors)
        except (UnicodeDecodeError, UnicodeError) as e:
            self.logger.warning(f"utf-8 解碼失敗: {e}，嘗試使用 latin-1")
            return self.decode_latin1(data, errors)
    
    def decode_latin1(self, data: Union[bytes, str], errors: str = 'ignore') -> str:
        """
        解碼 latin-1 編碼的數據（最寬鬆的編碼，不會失敗）
        
        Args:
            data: 要解碼的數據（bytes 或 str）
            errors: 錯誤處理策略（'ignore', 'replace', 'strict'）
            
        Returns:
            解碼後的字符串
        """
        if isinstance(data, str):
            return data
        
        try:
            return data.decode('latin-1', errors=errors)
        except Exception as e:
            self.logger.error(f"latin-1 解碼失敗: {e}")
            return data.decode('latin-1', errors='replace')
    
    def decode_big5(self, data: Union[bytes, str], errors: str = 'ignore') -> str:
        """
        解碼 big5 編碼的數據（繁體中文）
        
        Args:
            data: 要解碼的數據（bytes 或 str）
            errors: 錯誤處理策略（'ignore', 'replace', 'strict'）
            
        Returns:
            解碼後的字符串
        """
        if isinstance(data, str):
            return data
        
        try:
            return data.decode('big5', errors=errors)
        except (UnicodeDecodeError, UnicodeError) as e:
            self.logger.warning(f"big5 解碼失敗: {e}，嘗試使用 utf-8")
            return self.decode_utf8(data, errors)
    
    def decode_auto(self, data: Union[bytes, str], preferred_encodings: Optional[List[str]] = None, errors: str = 'ignore') -> str:
        """
        自動嘗試多種編碼解碼數據
        
        Args:
            data: 要解碼的數據（bytes 或 str）
            preferred_encodings: 優先嘗試的編碼列表（如果為 None，使用默認列表）
            errors: 錯誤處理策略（'ignore', 'replace', 'strict'）
            
        Returns:
            解碼後的字符串
        """
        if isinstance(data, str):
            return data
        
        if preferred_encodings is None:
            preferred_encodings = self.chinese_encodings + self.common_encodings
            # 去重並保持順序
            seen = set()
            preferred_encodings = [x for x in preferred_encodings if not (x in seen or seen.add(x))]
        
        for encoding in preferred_encodings:
            try:
                decoded = data.decode(encoding, errors=errors)
                self.logger.debug(f"成功使用 {encoding} 編碼解碼數據")
                return decoded
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        # 如果所有編碼都失敗，使用 latin-1（最寬鬆的編碼）
        self.logger.warning("所有編碼嘗試失敗，使用 latin-1 作為最後手段")
        return self.decode_latin1(data, errors)
    
    def safe_decode_string(self, value: Any, default: str = '') -> str:
        """
        安全地解碼字符串值，處理各種類型的輸入
        
        Args:
            value: 要解碼的值（可能是 bytes, str, 或其他類型）
            default: 解碼失敗時的默認值
            
        Returns:
            解碼後的字符串
        """
        if value is None:
            return default
        
        if isinstance(value, str):
            # 檢查字符串是否包含無法編碼的字符
            try:
                value.encode('utf-8')
                return value
            except UnicodeEncodeError:
                # 如果字符串包含無法編碼的字符，嘗試重新編碼
                try:
                    return value.encode('gb2312', errors='ignore').decode('gbk', errors='ignore')
                except:
                    return value.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
        
        if isinstance(value, bytes):
            return self.decode_auto(value)
        
        # 對於其他類型，嘗試轉換為字符串
        try:
            return str(value)
        except Exception as e:
            self.logger.warning(f"無法轉換值為字符串: {e}")
            return default
    
    def decode_windows1258(self, data: Union[bytes, str], errors: str = 'ignore') -> str:
        """
        解碼 Windows-1258 編碼的數據（越南文）
        
        Args:
            data: 要解碼的數據（bytes 或 str）
            errors: 錯誤處理策略（'ignore', 'replace', 'strict'）
            
        Returns:
            解碼後的字符串
        """
        if isinstance(data, str):
            return data
        
        try:
            return data.decode('windows-1258', errors=errors)
        except (UnicodeDecodeError, UnicodeError) as e:
            self.logger.warning(f"windows-1258 解碼失敗: {e}，嘗試使用 utf-8")
            return self.decode_utf8(data, errors)
    
    def is_corrupted(self, text: str) -> bool:
        """
        檢測字串是否為亂碼
        
        Args:
            text: 要檢測的字串
        
        Returns:
            如果是亂碼返回 True，否則返回 False
        """
        if not isinstance(text, str):
            return False
        
        # 檢查亂碼特徵
        # 1. 包含替換字元
        if "�" in text:
            return True
        
        # 2. 包含過多的問號（可能是解碼失敗的結果）
        if text.count("?") > len(text) * 0.1:  # 超過 10% 是問號
            return True
        
        # 3. 檢查是否包含不合理的字元組合
        # 例如：連續的問號、問號與其他字元的異常組合
        if re.search(r'\?{2,}', text):  # 連續兩個以上的問號
            return True
        
        return False
    
    def try_fix_mojibake(self, text: str, common_chars: str = '的是一有我不人這中大為上來到和說們印表機問題') -> str:
        """
        嘗試修正 Big5/UTF-8 誤解碼產生的亂碼。
        僅在修正後「常見字更多且罕見字更少」時才採用，避免變成另一種亂碼。
        """
        if not text or not isinstance(text, str):
            return text
        # 罕見字範圍：CJK Ext A/B 等，用於判斷是否變更差
        def _rare_count(s):
            return sum(1 for c in s if '\u3400' <= c <= '\u4dbf' or '\u4e00' <= c <= '\u9fff')
        def _common_score(s):
            return sum(1 for c in common_chars if c in s)
        def _has_too_many_rare(s, threshold=0.15):
            r = _rare_count(s)
            return r > 20 and (r / max(len(s), 1)) > threshold
        best, best_score = text, _common_score(text)
        orig_rare = _rare_count(text)
        for name, fn in [
            ('utf8_as_big5', lambda s: s.encode('big5', errors='ignore').decode('utf-8', errors='ignore')),
            ('big5_as_utf8', lambda s: s.encode('utf-8', errors='ignore').decode('big5', errors='ignore')),
        ]:
            try:
                fixed = fn(text)
                if len(fixed) < len(text) * 0.5 or self.is_corrupted(fixed):
                    continue
                sc = _common_score(fixed)
                fixed_rare = _rare_count(fixed)
                if sc > best_score and fixed_rare <= orig_rare * 1.2:
                    best, best_score = fixed, sc
                    self.logger.debug(f"mojibake 修正成功 ({name})")
            except (UnicodeDecodeError, UnicodeEncodeError):
                continue
        return best
    
    def strip_rare_cjk(self, text: str, keep_chars: str = '') -> str:
        """
        移除罕見 CJK 字元（可能為亂碼）。保留 ASCII、常用繁中、及 keep_chars。
        """
        if not text or not isinstance(text, str):
            return text
        common = set('的一是在有不人這中大為上來到和說們印表機問題端該實設使常追細以就帳豐親愛先說種方列路高加座出現卡住抓刪除目擊台確誤時因類似基全跑透程用集身定可才果例能') | set(keep_chars) | set(chr(i) for i in range(32, 127))
        def _ok(c):
            if c in common or c in ' \n\r\t':
                return True
            if '\u3400' <= c <= '\u4dbf':  # CJK Ext A，多為罕見字
                return False
            if '\u4e00' <= c <= '\u9fff':
                return c in common
            return True
        return ''.join(c for c in text if _ok(c))
    
    def detect_encoding(self, data: bytes) -> str:
        """
        自動檢測 bytes 的編碼
        
        Args:
            data: 要檢測的 bytes 資料
        
        Returns:
            檢測到的編碼名稱，如果無法檢測則返回 'utf-8'
        """
        if not isinstance(data, (bytes, bytearray)):
            return "utf-8"
        
        # 使用 chardet 自動檢測
        if HAS_CHARDET:
            try:
                result = chardet.detect(data)
                if result and result.get("encoding") and result.get("confidence", 0) > 0.7:
                    detected_encoding = result["encoding"].lower()
                    # 將檢測到的編碼映射到標準名稱
                    encoding_map = {
                        "gb2312": "gb2312",
                        "gbk": "gbk",
                        "big5": "big5",
                        "windows-1258": "windows-1258",
                        "windows-1252": "windows-1252",
                        "iso-8859-1": "iso-8859-1",
                        "latin-1": "iso-8859-1",
                        "cp1252": "windows-1252",
                        "cp1258": "windows-1258",
                    }
                    return encoding_map.get(detected_encoding, detected_encoding)
            except Exception:
                pass
        
        return "utf-8"
    
    def fix_encoding(self, text: str) -> str:
        """
        嘗試修復錯誤編碼的字串
        
        如果字串已經是錯誤編碼的字串（例如越南文被錯誤解碼），
        嘗試將其重新編碼為 bytes，然後用正確編碼解碼。
        
        Args:
            text: 要修復的字串
        
        Returns:
            修復後的字串，如果無法修復則返回原字串
        """
        if not isinstance(text, str):
            return str(text)
        
        # 如果字串看起來正常，直接返回
        if not self.is_corrupted(text):
            return text
        
        # 嘗試將字串重新編碼為 bytes，然後用不同編碼解碼
        # 常見情況：字串是用錯誤編碼解碼的，需要重新編碼回 bytes
        
        # 先嘗試用常見的錯誤編碼重新編碼
        error_encodings = ["latin1", "iso-8859-1", "windows-1252", "cp1252"]
        
        for error_enc in error_encodings:
            try:
                # 將字串用錯誤編碼重新編碼為 bytes
                bytes_data = text.encode(error_enc, errors="ignore")
                
                # 然後用正確的編碼解碼
                for correct_enc in self.supported_encodings:
                    if correct_enc == error_enc:
                        continue
                    try:
                        fixed = bytes_data.decode(correct_enc)
                        # 檢查修復後的結果是否更好
                        if not self.is_corrupted(fixed):
                            # 檢查是否包含越南文字元（表示可能修復成功）
                            has_vietnamese = any(ord(c) >= 0x0100 and ord(c) <= 0x1EF9 for c in fixed[:100])
                            if has_vietnamese or "?" not in fixed or fixed.count("?") < text.count("?"):
                                return fixed
                    except (UnicodeDecodeError, LookupError):
                        continue
            except (UnicodeEncodeError, LookupError):
                continue
        
        # 如果無法修復，返回原字串
        return text
    
    def decode_value(self, value: Any) -> str:
        """
        統一的解碼介面，自動處理 bytes 或字串
        
        Args:
            value: 要解碼的值（可能是 bytes、字串或其他類型）
        
        Returns:
            解碼後的字串
        """
        if value is None:
            return ""
        
        # 如果是 bytes，直接解碼
        if isinstance(value, (bytes, bytearray)):
            return self.decode_bytes(value)
        
        # 如果是字串，檢查是否需要修復
        if isinstance(value, str):
            # 檢查是否為亂碼
            if self.is_corrupted(value):
                # 嘗試修復
                fixed = self.fix_encoding(value)
                return fixed
            return value
        
        # 其他類型轉換為字串
        return str(value)
    
    def decode_bytes(self, data: bytes) -> str:
        """
        解碼 bytes 資料，嘗試多種編碼（包含越南文支援）
        
        Args:
            data: 要解碼的 bytes 資料
        
        Returns:
            解碼後的字串
        """
        if not isinstance(data, (bytes, bytearray)):
            return str(data)
        
        # 先嘗試自動檢測
        detected_encoding = self.detect_encoding(data)
        if detected_encoding != "utf-8":
            # 將檢測到的編碼放在優先位置
            encodings = [detected_encoding] + [e for e in self.supported_encodings if e != detected_encoding]
        else:
            encodings = self.supported_encodings
        
        # 嘗試各種編碼
        for encoding in encodings:
            try:
                decoded = data.decode(encoding)
                # 檢查解碼結果是否包含替換字元
                if "�" not in decoded:
                    # 進一步檢查是否為亂碼
                    if not self.is_corrupted(decoded):
                        return decoded
            except (UnicodeDecodeError, LookupError):
                continue
        
        # 如果都失敗，使用 UTF-8 並保留錯誤字元
        try:
            return data.decode("utf-8", errors="replace")
        except Exception:
            return str(data)
    
    def fix_gb2312_error(self, data: Union[bytes, str], fallback_encodings: Optional[List[str]] = None) -> str:
        """
        修復 gb2312 編碼錯誤，使用更寬鬆的編碼
        
        這個方法專門用於處理 extract_msg 等庫中出現的 gb2312 編碼錯誤
        
        Args:
            data: 要處理的數據（bytes 或 str）
            fallback_encodings: 回退編碼列表（默認為 ['gbk', 'gb18030', 'utf-8']）
            
        Returns:
            修復後的字符串
        """
        if isinstance(data, str):
            # 如果已經是字符串，嘗試檢測並修復編碼問題
            try:
                # 嘗試重新編碼為 gb2312 再解碼為 gbk
                return data.encode('gb2312', errors='ignore').decode('gbk', errors='ignore')
            except:
                return data
        
        if fallback_encodings is None:
            fallback_encodings = ['big5', 'gbk', 'gb18030', 'utf-8']
        
        # 嘗試多種編碼
        for encoding in fallback_encodings:
            try:
                decoded = data.decode(encoding, errors='ignore')
                self.logger.debug(f"使用 {encoding} 編碼修復 gb2312 錯誤")
                return decoded
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        # 如果都失敗，使用 latin-1
        return self.decode_latin1(data, errors='ignore')
    
    def register_codec_error_handler(self, handler_name: str = 'gbk_fallback'):
        """
        註冊編碼錯誤處理器到 codecs 模組
        
        Args:
            handler_name: 處理器名稱
        """
        import codecs
        
        def gbk_fallback_handler(error):
            """將 gb2312 錯誤轉換為 gbk 解碼"""
            try:
                decoded, consumed = codecs.codecs_decode(
                    error.object[error.start:error.end], 
                    'gbk', 
                    error
                )
                return decoded, consumed
            except:
                try:
                    decoded, consumed = codecs.codecs_decode(
                        error.object[error.start:error.end], 
                        'gb18030', 
                        error
                    )
                    return decoded, consumed
                except:
                    return ('?', error.end)
        
        try:
            codecs.register_error(handler_name, gbk_fallback_handler)
            self.logger.debug(f"成功註冊編碼錯誤處理器: {handler_name}")
            return True
        except Exception as e:
            self.logger.warning(f"註冊編碼錯誤處理器失敗: {e}")
            return False
    
    def create_msg_with_encoding(self, file_path: str, msg_class) -> Any:
        """
        創建 MSG 文件對象，處理編碼問題
        
        這個方法統一處理 extract_msg.Message 的創建過程，自動處理編碼錯誤
        
        Args:
            file_path: MSG 文件路徑
            msg_class: extract_msg.Message 類別
            
        Returns:
            MSG 對象實例
            
        Raises:
            Exception: 如果所有編碼嘗試都失敗
        """
        msg = None
        
        # 預先註冊編碼錯誤處理器（預防性措施）
        self.register_codec_error_handler('gbk_fallback')
        
        # 方法 1: 嘗試使用 extract_msg 的編碼參數（如果支持）
        try:
            import inspect
            sig = inspect.signature(msg_class.__init__)
            if 'overrideEncoding' in sig.parameters:
                # 優先使用 chardet 自動偵測；否則依序嘗試並驗證 body 非亂碼
                encodings_to_try = ['chardet'] if HAS_CHARDET else []
                encodings_to_try += ['utf-8', 'big5', 'gbk', 'gb18030']
                for encoding in encodings_to_try:
                    msg = None
                    try:
                        msg = msg_class(file_path, overrideEncoding=encoding)
                        body = getattr(msg, 'body', None) or ''
                        html = getattr(msg, 'htmlBody', None) or ''
                        sample = (str(body) + str(html))[:2000]
                        if sample and self.is_corrupted(sample):
                            self.logger.debug(f"使用 {encoding} 編碼解析後 body 仍為亂碼，嘗試下一編碼")
                            if hasattr(msg, 'close'):
                                msg.close()
                            continue
                        self.logger.debug(f"使用 {encoding} 編碼成功解析 MSG 文件（body 無亂碼）")
                        return msg
                    except (UnicodeDecodeError, UnicodeError, TypeError, ValueError) as e:
                        self.logger.debug(f"使用 {encoding} 編碼失敗: {e}")
                        if msg and hasattr(msg, 'close'):
                            try:
                                msg.close()
                            except Exception:
                                pass
                        continue
        except Exception as e:
            self.logger.debug(f"檢查 overrideEncoding 參數時發生錯誤: {e}")
        
        # 方法 2: 如果不支持 overrideEncoding 或方法 1 失敗，使用默認方式
        try:
            msg = msg_class(file_path)
            return msg
        except Exception as decode_error:
            # 捕獲所有異常，因為編碼錯誤可能被包裝在其他異常中
            error_msg = str(decode_error).lower()
            error_type = type(decode_error).__name__
            
            # 檢查是否為編碼相關錯誤
            is_encoding_error = (
                isinstance(decode_error, (UnicodeDecodeError, UnicodeError)) or
                'gb2312' in error_msg or 
                'codec' in error_msg or
                'decode' in error_msg or
                'encoding' in error_msg
            )
            
            if is_encoding_error:
                # 方法 3: 如果出現編碼錯誤，嘗試使用 monkey patch 修改編碼行為
                self.logger.warning(f"檢測到編碼錯誤 ({error_type}): {decode_error}，嘗試使用 monkey patch 修復")
                
                # 嘗試 monkey patch extract_msg 內部的編碼設置
                try:
                    # 檢查 extract_msg 是否有編碼相關的模組
                    import extract_msg
                    if hasattr(extract_msg, 'constants'):
                        # 嘗試修改默認編碼
                        original_encoding = getattr(extract_msg.constants, 'DEFAULT_ENCODING', None)
                        if original_encoding == 'gb2312':
                            extract_msg.constants.DEFAULT_ENCODING = 'gbk'
                            self.logger.debug("已將 extract_msg 默認編碼從 gb2312 改為 gbk")
                    elif hasattr(extract_msg, 'DEFAULT_ENCODING'):
                        if extract_msg.DEFAULT_ENCODING == 'gb2312':
                            extract_msg.DEFAULT_ENCODING = 'gbk'
                            self.logger.debug("已將 extract_msg 默認編碼從 gb2312 改為 gbk")
                except Exception as patch_error:
                    self.logger.debug(f"Monkey patch 失敗: {patch_error}")
                
                # 重新嘗試創建 Message 對象
                try:
                    msg = msg_class(file_path)
                    return msg
                except Exception as e2:
                    self.logger.warning(f"編碼錯誤處理後仍失敗: {e2}")
                    # 最後嘗試：使用錯誤處理策略直接讀取文件並手動處理
                    try:
                        return self._create_msg_with_fallback(file_path, msg_class)
                    except Exception as e3:
                        self.logger.error(f"所有編碼處理方法都失敗: {e3}")
                        raise decode_error  # 重新拋出原始錯誤
            else:
                # 其他類型的錯誤，直接重新拋出
                raise
    
    def _create_msg_with_fallback(self, file_path: str, msg_class) -> Any:
        """
        使用回退方法創建 MSG 對象
        
        當標準方法都失敗時，嘗試使用更底層的方法
        """
        # 嘗試直接讀取文件並使用 gbk 編碼處理
        try:
            # 如果 extract_msg 支持從 bytes 創建，嘗試這種方式
            with open(file_path, 'rb') as f:
                file_data = f.read()
            
            # 嘗試使用臨時文件並設置編碼環境變數
            import tempfile
            import os as os_module
            
            # 設置環境變數（可能對某些庫有效）
            original_lang = os_module.environ.get('LANG', None)
            original_lc_all = os_module.environ.get('LC_ALL', None)
            
            try:
                # 設置編碼相關環境變數
                os_module.environ['LANG'] = 'zh_CN.GBK'
                os_module.environ['LC_ALL'] = 'zh_CN.GBK'
                
                # 重新嘗試創建
                msg = msg_class(file_path)
                return msg
            finally:
                # 恢復環境變數
                if original_lang is not None:
                    os_module.environ['LANG'] = original_lang
                elif 'LANG' in os_module.environ:
                    del os_module.environ['LANG']
                    
                if original_lc_all is not None:
                    os_module.environ['LC_ALL'] = original_lc_all
                elif 'LC_ALL' in os_module.environ:
                    del os_module.environ['LC_ALL']
        except Exception as e:
            self.logger.debug(f"回退方法失敗: {e}")
            raise
    
    def extract_msg_attributes(self, msg_obj: Any) -> Dict[str, Any]:
        """
        統一提取 MSG 對象的所有屬性，並進行編碼處理
        
        這個方法會捕獲訪問屬性時可能發生的編碼錯誤（例如 extract_msg 內部使用 gb2312 解碼失敗）
        
        Args:
            msg_obj: extract_msg.Message 對象
            
        Returns:
            包含所有屬性的字典，所有字符串值都已正確解碼
        """
        result = {}
        
        # 在提取屬性前，monkey patch extract_msg.utils.decodeRfc2047
        # 這需要在訪問任何屬性之前完成，因為錯誤發生在屬性訪問時
        original_decodeRfc2047 = None
        patched_decodeRfc2047 = None
        
        try:
            import extract_msg.utils as msg_utils
            if hasattr(msg_utils, 'decodeRfc2047'):
                original_decodeRfc2047 = msg_utils.decodeRfc2047
                
                # 創建包裝函數，使用 gbk 替代 gb2312
                def patched_decode(s):
                    try:
                        # 先嘗試原始方法
                        return original_decodeRfc2047(s)
                    except (UnicodeDecodeError, UnicodeError) as e:
                        # 如果原始解碼失敗（通常是 gb2312 錯誤），嘗試使用 gbk
                        self.logger.debug(f"decodeRfc2047 使用 gb2312 失敗，嘗試使用 gbk: {e}")
                        if isinstance(s, bytes):
                            try:
                                return s.decode('big5', errors='ignore')
                            except:
                                try:
                                    return s.decode('gbk', errors='ignore')
                                except:
                                    try:
                                        return s.decode('gb18030', errors='ignore')
                                    except:
                                        return s.decode('utf-8', errors='ignore')
                        elif isinstance(s, str):
                            # 如果已經是字符串，嘗試重新編碼
                            try:
                                # 先嘗試 big5 重新編碼
                                return s.encode('big5', errors='ignore').decode('big5', errors='ignore')
                            except:
                                try:
                                    return s.encode('gb2312', errors='ignore').decode('gbk', errors='ignore')
                                except:
                                    return s
                        return s
                
                patched_decodeRfc2047 = patched_decode
                msg_utils.decodeRfc2047 = patched_decode
                self.logger.debug("已 monkey patch extract_msg.utils.decodeRfc2047（在提取屬性前）")
        except Exception as e:
            self.logger.debug(f"嘗試 monkey patch decodeRfc2047 時發生錯誤: {e}")
        
        def safe_get_attr_with_encoding_fix(obj, attr_name: str, default: str = ''):
            """
            安全地獲取屬性，處理 extract_msg 內部編碼錯誤
            
            當 extract_msg 內部使用 gb2312 解碼失敗時，會捕獲錯誤並嘗試修復
            關鍵：需要在訪問屬性時就捕獲錯誤，因為錯誤發生在屬性訪問過程中
            """
            try:
                # 嘗試獲取屬性（這裡可能會觸發 extract_msg 內部的編碼錯誤）
                value = getattr(obj, attr_name, None)
                if value is None:
                    return default
                # 使用 safe_decode_string 處理編碼
                return self.safe_decode_string(value, default)
            except (UnicodeDecodeError, UnicodeError) as e:
                # 捕獲編碼錯誤（發生在 extract_msg 內部，例如 utils.decodeRfc2047）
                # 注意：monkey patch 應該已經在函數開始時應用，如果仍然失敗，嘗試其他方法
                self.logger.warning(f"訪問 {attr_name} 屬性時發生編碼錯誤（monkey patch 後仍失敗）: {e}")
                try:
                    # 嘗試直接訪問底層屬性並手動處理編碼
                    if hasattr(obj, '_raw_' + attr_name):
                        raw_value = getattr(obj, '_raw_' + attr_name, None)
                        if raw_value:
                            return self.fix_gb2312_error(raw_value if isinstance(raw_value, bytes) else str(raw_value).encode('latin-1', errors='ignore'))
                    
                    # 如果無法獲取原始數據，返回默認值
                    return default
                except Exception as e2:
                    self.logger.debug(f"修復 {attr_name} 編碼錯誤時發生錯誤: {e2}")
                    return default
            except Exception as e:
                # 其他類型的錯誤
                self.logger.debug(f"訪問 {attr_name} 屬性時發生錯誤: {e}")
                return default
        
        # 提取基本郵件屬性（使用安全方法）
        result['subject'] = safe_get_attr_with_encoding_fix(msg_obj, 'subject', '')
        result['sender'] = safe_get_attr_with_encoding_fix(msg_obj, 'sender', '')
        result['body'] = safe_get_attr_with_encoding_fix(msg_obj, 'body', '')
        
        # 提取日期
        try:
            date_value = getattr(msg_obj, 'date', None)
            result['date'] = self.safe_decode_string(date_value, '') if date_value else ''
        except (UnicodeDecodeError, UnicodeError) as e:
            self.logger.warning(f"訪問 date 屬性時發生編碼錯誤: {e}")
            result['date'] = ''
        
        # 提取收件者資訊
        result['to'] = safe_get_attr_with_encoding_fix(msg_obj, 'to', '')
        result['cc'] = safe_get_attr_with_encoding_fix(msg_obj, 'cc', '')
        
        # 提取附件資訊
        try:
            attachments = getattr(msg_obj, 'attachments', []) or []
        except (UnicodeDecodeError, UnicodeError) as e:
            self.logger.warning(f"訪問 attachments 屬性時發生編碼錯誤: {e}")
            attachments = []
        
        result['attachments'] = []
        for att in attachments:
            if att:
                try:
                    att_info = {
                        'name': safe_get_attr_with_encoding_fix(
                            att, 
                            'name', 
                            safe_get_attr_with_encoding_fix(att, 'shortFilename', '')
                        ),
                        'longFilename': safe_get_attr_with_encoding_fix(
                            att, 
                            'longFilename', 
                            safe_get_attr_with_encoding_fix(att, 'shortFilename', '')
                        ),
                        'size': getattr(att, 'dataLength', 0) if hasattr(att, 'dataLength') else 0
                    }
                    result['attachments'].append(att_info)
                except Exception as e:
                    self.logger.debug(f"處理附件時發生錯誤: {e}")
                    # 跳過這個附件，繼續處理下一個
                    continue
        
        # 恢復原始的 decodeRfc2047 函數
        if original_decodeRfc2047 is not None and patched_decodeRfc2047 is not None:
            try:
                import extract_msg.utils as msg_utils
                msg_utils.decodeRfc2047 = original_decodeRfc2047
                self.logger.debug("已恢復 extract_msg.utils.decodeRfc2047")
            except Exception as e:
                self.logger.debug(f"恢復 decodeRfc2047 時發生錯誤: {e}")
        
        return result


# 預編譯正則表達式以提高效率
_MARKDOWN_CODE_PATTERN = re.compile(r'```(?:json)?\s*(\[.*?\])\s*```', re.DOTALL)
_QUOTED_KEYWORD_PATTERN = re.compile(r'["\']([^"\']{3,10})["\']')

# 批次處理大小限制（避免一次性發送過多請求）
# 注意：TextProcessor /chat/batch API 預設限制每批最多 100 個問句
# 但可以通過 max_batch_size 參數請求更大的批次（例如 200）
_BATCH_SIZE_LIMIT = 190  # 每批最多處理 190 個 prompts（留出安全邊際，服務器端允許 200）


def batch_chat_api(
        prompts: List[str],
        base_url: str = None,
        provider: str = None,
        model: str = None,
        system_prompt: str = '',
        max_tokens: int = 200,
        temperature: float = 0.3,
        batch_size: int = None,
        timeout: int = 600,
        logger: logging.Logger = None
    ) -> List[Dict[str, Any]]:
    """
    批次調用 TextProcessor 的 /chat/batch API，使用並行處理
    
    注意：/chat/batch API 不支援 system_prompt 參數，如果需要 system_prompt，
    應該在調用此函數前將其整合到每個 prompt 中。
    
    Args:
        prompts: 提示詞列表
        base_url: TextProcessor API 的 base URL（如果為 None，從配置讀取）
        provider: LLM 提供者（如果為 None，從配置讀取）
        model: 模型名稱（如果為 None，從配置讀取）
        system_prompt: 系統提示詞（注意：batch API 不支援，此參數保留以保持介面一致性）
        max_tokens: 最大 token 數
        temperature: 溫度參數
        batch_size: 每批處理的大小（如果為 None，使用 _BATCH_SIZE_LIMIT）
        timeout: 請求超時時間（秒）
        logger: 日誌記錄器（可選）
        
    Returns:
        結果列表，每個元素是一個字典：
        - 'success': bool - 是否成功
        - 'output': str - LLM 回應文本（成功時）
        - 'error': str - 錯誤訊息（失敗時）
        - 'original_index': int - 原始索引（用於映射回原始列表）
    """
    if not prompts:
        return []
    
    if batch_size is None:
        batch_size = _BATCH_SIZE_LIMIT
    else:
        # 確保不超過 API 限制（最多 200 個，通過 max_batch_size 參數請求）
        if batch_size > 200:
            logger.warning(f"批次大小 {batch_size} 超過 API 限制（200），自動調整為 190")
            batch_size = 190
    
    # 從配置讀取默認值（如果未提供）
    if base_url is None or provider is None or model is None:
        llm_config = m_config.get('llm', {}) if isinstance(m_config, dict) else {}
        external_config = m_config.get('external', {}) if isinstance(m_config, dict) else {}
        textprocessor_config = external_config.get('textprocessor', {})
        
        if base_url is None:
            base_url = llm_config.get('base_url') or textprocessor_config.get('base_url', 'http://10.1.3.127:7017')
        if provider is None:
            provider = llm_config.get('chat_provider') or textprocessor_config.get('chat_provider', 'remote')
        if model is None:
            model = llm_config.get('chat_model') or textprocessor_config.get('chat_model', 'remote8b')
    
    batch_chat_url = f"{base_url}/chat/batch"
    
    # 初始化結果列表
    results = []
    total_success = 0
    total_errors = 0
    
    # 分批處理，避免一次性發送過多請求
    total_batches = (len(prompts) + batch_size - 1) // batch_size
    
    for batch_idx in range(total_batches):
        start_idx = batch_idx * batch_size
        end_idx = min(start_idx + batch_size, len(prompts))
        batch_prompts = prompts[start_idx:end_idx]
        
        try:
            # 注意：/chat/batch API 不支援 system_prompt 參數
            batch_chat_payload = {
                "prompts": batch_prompts,
                "provider": provider,
                "model": model,
                "max_tokens": max_tokens,
                "temperature": temperature,
                "parallel": True,  # 啟用並行處理
                "max_batch_size": 200  # 請求服務器允許更大的批次大小
            }
            
            # 記錄請求詳情（僅第一個批次，避免日誌過多）
            if batch_idx == 0:
                logger.info(f"批次 chat API 請求詳情:")
                logger.info(f"  URL: {batch_chat_url}")
                logger.info(f"  Provider: {provider}")
                logger.info(f"  Model: {model}")
                logger.info(f"  Max tokens: {max_tokens}")
                logger.info(f"  Temperature: {temperature}")
                logger.info(f"  Parallel: True")
                logger.info(f"  批次大小: {len(batch_prompts)}")
                logger.info(f"  第一個 prompt 長度: {len(batch_prompts[0]) if batch_prompts else 0} 字元")
                logger.info(f"  第一個 prompt 預覽: {batch_prompts[0][:200] if batch_prompts else 'N/A'}...")
                # 記錄 payload 的鍵（不記錄完整內容，避免日誌過大）
                logger.debug(f"  Payload keys: {list(batch_chat_payload.keys())}")
                logger.debug(f"  Prompts 類型: {type(batch_prompts)}, 長度: {len(batch_prompts)}")
                # 檢查 prompts 是否為列表且每個元素都是字符串
                if batch_prompts:
                    logger.debug(f"  第一個 prompt 類型: {type(batch_prompts[0])}")
                    if not isinstance(batch_prompts[0], str):
                        logger.error(f"  錯誤：第一個 prompt 不是字符串類型！類型: {type(batch_prompts[0])}")
            
            response = requests.post(batch_chat_url, json=batch_chat_payload, timeout=timeout)
            
            # 記錄響應狀態（即使失敗也要記錄）
            if batch_idx == 0:
                logger.info(f"  響應狀態碼: {response.status_code}")
                logger.info(f"  響應 headers: {dict(response.headers)}")
            
            # 嘗試解析響應內容（即使狀態碼不是 200）
            try:
                response_text = response.text
                if batch_idx == 0:
                    logger.info(f"  響應內容預覽: {response_text[:500]}...")
            except Exception as e:
                logger.warning(f"  無法讀取響應內容: {e}")
            
            response.raise_for_status()
            result = response.json()
            
            # 解析批次響應結果
            batch_results = result.get('results', [])
            
            # 處理每個結果
            for result_idx, item in enumerate(batch_results):
                original_idx = start_idx + result_idx
                
                if item.get('error'):
                    total_errors += 1
                    results.append({
                        'success': False,
                        'error': item.get('error'),
                        'output': '',
                        'original_index': original_idx
                    })
                    if total_errors <= 5:  # 只記錄前 5 個錯誤
                        logger.warning(f"批次 chat API 失敗（索引 {original_idx}）: {item.get('error')}")
                else:
                    result_data = item.get('result', {})
                    output = result_data.get('output', '').strip()
                    
                    # 清理響應文本
                    output = output.replace("<|eot_id|>", "").strip()
                    
                    results.append({
                        'success': True,
                        'output': output,
                        'error': '',
                        'original_index': original_idx
                    })
                    total_success += 1
            
            # 顯示批次進度
            if total_batches > 1:
                logger.debug(f"批次處理進度: {batch_idx + 1}/{total_batches} (已處理 {end_idx}/{len(prompts)} 個項目)")
                
        except requests.exceptions.HTTPError as e:
            total_errors += len(batch_prompts)
            error_msg = str(e)
            
            # 嘗試獲取響應內容以獲取更詳細的錯誤信息
            try:
                if hasattr(e, 'response') and e.response is not None:
                    error_detail = e.response.text
                    logger.error(f"批次 {batch_idx + 1}/{total_batches} HTTP 錯誤詳情:")
                    logger.error(f"  狀態碼: {e.response.status_code}")
                    logger.error(f"  錯誤訊息: {error_msg}")
                    logger.error(f"  響應內容: {error_detail[:1000]}")  # 限制長度避免日誌過大
                    
                    # 如果是第一個批次，記錄完整的 payload 結構（用於調試）
                    if batch_idx == 0:
                        logger.error(f"  發送的 payload 結構:")
                        logger.error(f"    prompts 數量: {len(batch_prompts)}")
                        logger.error(f"    prompts 類型: {type(batch_prompts)}")
                        if batch_prompts:
                            logger.error(f"    第一個 prompt 類型: {type(batch_prompts[0])}")
                            logger.error(f"    第一個 prompt 長度: {len(str(batch_prompts[0]))}")
                            logger.error(f"    第一個 prompt 內容: {str(batch_prompts[0])[:500]}")
                        logger.error(f"    provider: {provider} (類型: {type(provider)})")
                        logger.error(f"    model: {model} (類型: {type(model)})")
                        logger.error(f"    max_tokens: {max_tokens} (類型: {type(max_tokens)})")
                        logger.error(f"    temperature: {temperature} (類型: {type(temperature)})")
            except Exception as parse_error:
                logger.error(f"  無法解析錯誤響應: {parse_error}")
            
            import traceback
            if batch_idx < 3:  # 只打印前 3 個批次的詳細 traceback
                logger.error(f"  完整 traceback:")
                traceback.print_exc()
            
            # 為失敗的批次添加錯誤結果
            for idx in range(start_idx, end_idx):
                results.append({
                    'success': False,
                    'error': error_msg,
                    'output': '',
                    'original_index': idx
                })
            continue
        except Exception as e:
            total_errors += len(batch_prompts)
            logger.error(f"批次 {batch_idx + 1}/{total_batches} 處理失敗: {e}")
            import traceback
            if batch_idx < 3:  # 只打印前 3 個批次的詳細 traceback
                traceback.print_exc()
            
            # 為失敗的批次添加錯誤結果
            for idx in range(start_idx, end_idx):
                results.append({
                    'success': False,
                    'error': str(e),
                    'output': '',
                    'original_index': idx
                })
            continue
    
    # 顯示統計資訊
    if total_errors > 0:
        logger.info(f"批次 chat API 統計: 成功 {total_success}, 失敗 {total_errors}")
    
    # 按 original_index 排序，確保順序正確
    results.sort(key=lambda x: x['original_index'])
    
    return results


def _generate_base_prompt_with_llm(
        content: str,
        heading: str = '',
        tags: List[str] = None,
        os_tags: List[str] = None,
        struc_tags: List[str] = None,
        contextual_tags: List[str] = None,
        content_tags: List[str] = None,
        llm_config: Dict[str, Any] = None
    ) -> str:
    """
    使用 LLM 根據文章結構給出信息摘要與用途，生成 base_prompt
    
    Args:
        content: 段落內容
        heading: 標題
        tags: 所有標籤列表
        os_tags: 作業系統/環境標籤
        struc_tags: 結構標籤
        contextual_tags: 脈絡標籤
        content_tags: 內容標籤
        llm_config: LLM 配置字典
        
    Returns:
        生成的 base_prompt 字符串，失敗時返回 fallback
    """
    try:
        # 載入 prompt 配置
        prompt_file = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'prompt', 'generate_base_prompt.json')
        if not os.path.exists(prompt_file):
            m_logger.warning(f"Base prompt 配置文件不存在: {prompt_file}，使用 fallback")
            return _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
        
        prompt_config = LOGger.load_json(prompt_file)
        if not prompt_config:
            m_logger.warning(f"無法讀取 base prompt 配置: {prompt_file}，使用 fallback")
            return _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
        
        # 設定 LLM 配置
        if llm_config is None:
            llm_config = m_config.get('llm', {}) if isinstance(m_config, dict) else {}
        
        external_config = m_config.get('external', {}) if isinstance(m_config, dict) else {}
        textprocessor_config = external_config.get('textprocessor', {})
        
        base_url = llm_config.get('base_url') or textprocessor_config.get('base_url', 'http://10.1.3.127:6017')
        chat_provider = llm_config.get('chat_provider') or textprocessor_config.get('chat_provider', 'remote')
        chat_model = llm_config.get('chat_model') or textprocessor_config.get('chat_model', 'remote8b')
        
        # 模型名稱轉換
        if chat_provider == 'openai':
            if chat_model == 'gpt-3.5-turbo':
                chat_model = 'gpt35_chat'
            elif chat_model == 'gpt-4o':
                chat_model = 'gpt4o_chat'
            elif chat_model == 'gpt-4o-mini':
                chat_model = 'o4_chat'
            elif chat_model == 'gpt-4':
                chat_model = 'gpt4_chat'
        
        # 檢查內容是否足夠（避免 LLM 返回錯誤消息）
        content_clean = content.strip() if content else ''
        min_content_length = 10  # 最少 10 個字符
        
        if len(content_clean) < min_content_length:
            m_logger.debug(f"內容過短 ({len(content_clean)} 字符)，直接使用 fallback")
            return _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
        
        # 準備輸入數據
        input_data = {
            'content': content[:500] if len(content) > 500 else content,  # 限制內容長度
            'heading': heading or '',
            'os_tags': os_tags or [],
            'struc_tags': struc_tags or [],
            'contextual_tags': contextual_tags or [],
            'content_tags': content_tags or [],
            'all_tags': tags or []
        }
        
        # 構建 prompt
        system_prompt = prompt_config.get('system_prompt', '')
        user_prompt_template = prompt_config.get('user_prompt_template', '')
        max_tokens = int(prompt_config.get('generation_config', {}).get('max_new_tokens', 100))
        temperature = float(prompt_config.get('generation_config', {}).get('temperature', 0.3))
        
        # 格式化用戶 prompt
        user_prompt = user_prompt_template.format(**input_data)
        
        url = f"{base_url}/chat"
        payload = {
            "prompt": user_prompt,
            "provider": chat_provider,
            "model": chat_model,
            "system_prompt": system_prompt,
            "max_tokens": max_tokens,
            "temperature": temperature
        }
        
        m_logger.debug(f"調用 TextProcessor 生成 base_prompt: {url}, provider={chat_provider}, model={chat_model}")
        
        timeout = llm_config.get('timeout', 30)
        # 直接使用 requests，避免循環引用
        response = requests.post(url, json=payload, timeout=timeout)
        response.raise_for_status()
        
        if response.status_code == 200:
            data = response.json()
            answer = data.get('output', '') or data.get('response', '') or ''
            
            if answer and answer.strip():
                # 清理回應（移除多餘的引號、換行等）
                base_prompt = answer.strip().strip('"\'').replace('\n', ' ').replace('\r', ' ')
                # 移除 LLM 結束標記 <|eot_id|>
                if '<|eot_id|>' in base_prompt:
                    base_prompt = base_prompt.split('<|eot_id|>')[0].strip()
                # 限制長度
                if len(base_prompt) > 200:
                    base_prompt = base_prompt[:200].strip()
                
                # 檢測錯誤消息模式（LLM 返回的錯誤提示）
                # 如果 base_prompt 包含錯誤消息模式，使用 fallback
                is_error_message = any(pattern in base_prompt for pattern in m_llm_error_patterns)
                if is_error_message:
                    m_logger.warning(f"LLM 返回錯誤消息，使用 fallback: {base_prompt[:50]}...")
                    return _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                
                m_logger.debug(f"LLM 生成 base_prompt 成功: {base_prompt[:50]}...")
                return base_prompt
            else:
                m_logger.warning("LLM 返回空的 base_prompt，使用 fallback")
                return _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
        else:
            m_logger.warning(f"TextProcessor base_prompt API 錯誤: {response.status_code}，使用 fallback")
            return _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
            
    except KeyboardInterrupt:
        m_logger.warning("生成 base_prompt 請求被用戶中斷，使用 fallback")
        return _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
    except Exception as e:
        m_logger.warning(f"生成 base_prompt 失敗: {e}，使用 fallback")
        return _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)

def _generate_fallback_base_prompt(
        heading: str = '',
        struc_tags: List[str] = None,
        contextual_tags: List[str] = None,
        content: str = ''
    ) -> str:
    """
    生成 fallback base_prompt（當 LLM 調用失敗時使用）
    
    Args:
        heading: 標題
        struc_tags: 結構標籤
        contextual_tags: 脈絡標籤
        content: 內容
        
    Returns:
        fallback base_prompt
    """
    # 優先使用 heading
    if heading and heading.strip():
        return heading.strip()
    
    # 如果沒有 heading，使用 struc_tags 組合
    if struc_tags:
        combined = ' | '.join(struc_tags[:3])  # 最多使用前3個
        if combined:
            return combined
    
    # 如果沒有 struc_tags，使用 contextual_tags 組合
    if contextual_tags:
        combined = ' | '.join(contextual_tags[:3])  # 最多使用前3個
        if combined:
            return combined
    
    # 最後 fallback：使用 content 的前50個字符
    if content:
        return content[:50].strip()
    
    return ''

def _is_system_storage_path(path: str) -> bool:
    path_sep = os.sep  # 使用系統路徑分隔符
    if not path:
        return False
    try:
        normalized = os.path.normcase(os.path.abspath(path)).rstrip(path_sep)
    except Exception:
        return False
    for root in m_storage_roots:
        if normalized == root or normalized.startswith(root + path_sep):
            return True
    return False

def segment_by_custom_pattern(text: str, sep: str) -> List[str]:
    """按自定義分段"""
    parts = text.split(sep)
    return [p.strip() for p in parts if p.strip()]

def segment_by_lines(text: str) -> List[str]:
    """按行分段"""
    return segment_by_custom_pattern(text, '\n')
    
def segment_by_paragraphs(text: str) -> List[str]:
    """按段落分段"""
    return segment_by_custom_pattern(text, '\n\n')

def split_text(text: str, max_length: int = 1000, min_length: int = 50) -> List[str]:
    """
    文字分段 - 改進版，支援多種分隔符
    
    Args:
        text: 原始文字
        max_length: 最大長度
        min_length: 最小長度（過濾過短的分段）
        
    Returns:
        分段後的文字列表
    """
    if not text or not text.strip():
        return []
    
    if len(text) <= max_length:
        return [text] if len(text) >= min_length else []
    
    chunks = []
    current_chunk = ""
    
    # 優先使用段落分隔（\n\n），然後是換行（\n），最後是句號（。）
    if '\n\n' in text:
        # 按段落分段
        paragraphs = text.split('\n\n')
        for para in paragraphs:
            para = para.strip()
            if not para:
                continue
            
            if len(current_chunk + para) <= max_length:
                if current_chunk:
                    current_chunk += '\n\n' + para
                else:
                    current_chunk = para
            else:
                if current_chunk:
                    chunk = current_chunk.strip()
                    if len(chunk) >= min_length:
                        chunks.append(chunk)
                # 如果單個段落超過 max_length，進一步分割
                if len(para) > max_length:
                    # 按句號分割長段落
                    sentences = para.split('。')
                    for sentence in sentences:
                        sentence = sentence.strip()
                        if not sentence:
                            continue
                        
                        if len(current_chunk + sentence) <= max_length:
                            if current_chunk:
                                current_chunk += '。' + sentence
                            else:
                                current_chunk = sentence
                        else:
                            if current_chunk:
                                chunk = current_chunk.strip()
                                if len(chunk) >= min_length:
                                    chunks.append(chunk)
                            current_chunk = sentence
                else:
                    current_chunk = para
    elif '\n' in text:
        # 按換行分段
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            if len(current_chunk + line) <= max_length:
                if current_chunk:
                    current_chunk += '\n' + line
                else:
                    current_chunk = line
            else:
                if current_chunk:
                    chunk = current_chunk.strip()
                    if len(chunk) >= min_length:
                        chunks.append(chunk)
                current_chunk = line
    else:
        # 按句號分段（原始邏輯）
        sentences = text.split('。')
        for sentence in sentences:
            sentence = sentence.strip()
            if not sentence:
                continue
            
            if len(current_chunk + sentence) <= max_length:
                if current_chunk:
                    current_chunk += '。' + sentence
                else:
                    current_chunk = sentence
            else:
                if current_chunk:
                    chunk = current_chunk.strip()
                    if len(chunk) >= min_length:
                        chunks.append(chunk)
                current_chunk = sentence
    
    if current_chunk:
        chunk = current_chunk.strip()
        if len(chunk) >= min_length:
            chunks.append(chunk)
    
    # 如果所有分段都被過濾掉（都太短），至少返回一個分段
    if not chunks and text.strip():
        return [text.strip()]
    
    return chunks

class ContextParser:
    """文件書寫模式檢測與分段策略解析器"""
    
    # 支援的模式類型
    PATTERN_CODE = "code"
    PATTERN_LOG = "log"
    PATTERN_COMMUNICATION = "communication"
    PATTERN_STRUCTURED_DOC = "structured_doc"
    PATTERN_PLAIN_TEXT = "plain_text"  # 預設模式
    
    def __init__(self, llm_service, config: Dict[str, Any]):
        """
        初始化 ContextParser
        
        Args:
            llm_service: LLM 服務實例（用於模式檢測）
            config: 系統配置字典
        """
        self.llm_service = llm_service
        self.config = config
        
        # 從 config.json 讀取 LLM 配置
        llm_config = config.get('llm', {})
        self.llm_timeout = llm_config.get('timeout', 600)
        segment_config = llm_config.get('segment', {})
        self.llm_temperature = segment_config.get('temperature', 0.3)
        self.llm_max_tokens = segment_config.get('max_prompt_length', 512)  # 使用 max_prompt_length 作為 max_tokens
        self.llm_provider = llm_config.get('chat_provider', 'remote')
        self.llm_model = llm_config.get('chat_model', 'remote8b')
        
        # 載入配置檔案
        self.base_path = Path(__file__).parent
        self.config_path = self.base_path / "config"
        self.logger = m_logger
        
        # 獲取 prompt 資料夾路徑（從 config.json 或使用預設路徑）
        # prompt 資料夾位於專案根目錄
        project_root = self.base_path.parent.parent
        prompt_path = config.get('prompt_path', None)
        if prompt_path:
            self.prompt_path = Path(prompt_path)
        else:
            # 預設使用專案根目錄下的 prompt 資料夾
            self.prompt_path = project_root / "prompt"
        
        # 載入模式檢測配置
        self.pattern_config = self._load_config("pattern_detection.json")
        
        # 載入分段策略配置
        self.code_config = self._load_config("code_segmentation.json")
        self.log_config = self._load_config("log_segmentation.json")
        self.communication_config = self._load_config("communication_segmentation.json")
        self.structured_doc_config = self._load_config("structured_doc_segmentation.json")
        
        # 載入日誌格式分析 prompt（用於 LLM 偵測時間戳格式和日誌級別）
        self.log_format_analysis_config = self._load_prompt_config("log_format_analysis.json")
        
        # 載入關鍵詞提取 prompt（用於內容標籤萃取）
        self.extract_keywords_config = self._load_prompt_config("extract_keywords.json")
        
        # 載入過濾口語用字 prompt（用於過濾 call_prompt 和 tags）
        self.filter_casual_terms_config = self._load_prompt_config("filter_casual_terms.json")
        # LLM 去重 multi_prompts 的 prompt 配置
        self.dedup_multi_prompts_config = self._load_prompt_config("dedup_multi_prompts.json")
        
        # 讀取過濾配置參數
        filter_config = llm_config.get('filter_casual_terms', {})
        self.filter_casual_terms_enabled = filter_config.get('enabled', True)
        self.filter_casual_terms_threshold = filter_config.get('threshold', 6)
        # multi_prompts 去重設定（LLM）
        dedup_cfg = (segment_config.get('dedup_multi_prompts') or {}) if isinstance(segment_config, dict) else {}
        self.dedup_multi_prompts_enabled = dedup_cfg.get('enabled', True)
        self.dedup_multi_prompts_min_count = int(dedup_cfg.get('min_count', 3))
        self.dedup_multi_prompts_batch_size = int(dedup_cfg.get('batch_size', 8))
        self.dedup_multi_prompts_max_items = int(dedup_cfg.get('max_items', 40))
        self.dedup_multi_prompts_max_keep = int(dedup_cfg.get('max_keep', 0))
        
        # 提取系統儲存區（collection）名稱，用於過濾無意義標籤
        self.system_collections = self._extract_system_collections(config)

        # 調試開關：環境變數 CTX_VERBOSE_LOG=1 或 config.debug.context_parser_verbose=True
        debug_cfg = (self.config.get('debug') or {}) if isinstance(self.config, dict) else {}
        env_verbose = os.environ.get('CTX_VERBOSE_LOG', '').lower() in ('1', 'true', 'yes')
        self.debug_verbose = bool(debug_cfg.get('context_parser_verbose')) or env_verbose

        # 並行執行器（可選）
        try:
            self.parallel_executor = ParallelLLMExecutor(llm_service, config) if llm_service and ParallelLLMExecutor else None
        except Exception as e:
            if self.logger:
                self.logger.warning(f"初始化 parallel_executor 失敗，改用串行: {e}")
            self.parallel_executor = None
    
    def _load_config(self, filename: str) -> Dict[str, Any]:
        """載入配置檔案（從 config 資料夾）"""
        try:
            config_file = self.config_path / filename
            if config_file.exists():
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                self.logger.warning(f"配置檔案不存在: {config_file}")
                return {}
        except Exception as e:
            LOGger.exception_process(e)
            self.logger.error(f"載入配置檔案失敗 {filename}: {e}")
            return {}
    
    def _load_prompt_config(self, filename: str) -> Dict[str, Any]:
        """載入 prompt 配置檔案（從 prompt 資料夾）"""
        try:
            prompt_file = self.prompt_path / filename
            if prompt_file.exists():
                with open(prompt_file, 'r', encoding='utf-8-sig') as f:
                # with open(prompt_file, 'r') as f:
                    return json.load(f)
            else:
                self.logger.warning(f"Prompt 檔案不存在: {prompt_file}")
                return {}
        except Exception as e:
            self.logger.error(f"載入 prompt 檔案失敗 {filename}: {e}")
            return {}
    
    def _filter_casual_terms(
            self, 
            terms: List[str], 
            max_result_count: Optional[int] = None
        ) -> List[str]:
        """
        過濾口語用字和無技術成分的用字（共用函數）
        
        Args:
            terms: 要過濾的詞列表
            max_result_count: 最大返回數量（可選）
            
        Returns:
            過濾後的詞列表
        """
        if not self.filter_casual_terms_enabled or not terms:
            return terms[:max_result_count] if max_result_count else terms
        
        filtered_terms = []
        terms_to_filter = []
        filter_prompts = []
        filter_indices = []
        
        # 1. 根據長度閾值分類
        for i, term in enumerate(terms):
            if len(term) <= self.filter_casual_terms_threshold:
                terms_to_filter.append(term)
                filter_prompt = self._build_filter_prompt(term)
                if filter_prompt:
                    filter_prompts.append(filter_prompt)
                    filter_indices.append(i)
            else:
                filtered_terms.append(term)
        
        # 2. 並行或串行判斷
        if filter_prompts and self.parallel_executor:
            # 並行模式
            filter_results = self.parallel_executor.batch_chat(filter_prompts)
            for (orig_idx, term), filter_result in zip(
                    [(filter_indices[i], terms_to_filter[i]) for i in range(len(filter_indices))], 
                    filter_results
                ):
                should_filter = self._parse_filter_result(filter_result) if filter_result else False
                if not should_filter:
                    filtered_terms.append(term)
        elif filter_prompts:
            # 串行模式
            for orig_idx, term in zip(filter_indices, terms_to_filter):
                filter_prompt = self._build_filter_prompt(term)
                if filter_prompt and self.llm_service:
                    try:
                        filter_result = self.llm_service.chat(
                            prompt=filter_prompt['prompt'],
                            system_prompt=filter_prompt.get('system_prompt', ''),
                            provider=filter_prompt.get('provider', self.llm_provider),
                            model=filter_prompt.get('model', self.llm_model),
                            max_tokens=filter_prompt.get('max_tokens', 50),
                            temperature=filter_prompt.get('temperature', 0.2)
                        )
                        should_filter = self._parse_filter_result(filter_result) if filter_result else False
                        if not should_filter:
                            filtered_terms.append(term)
                    except Exception as e:
                        self.logger.warning(f"過濾判斷失敗（串行）: {e}，保留該詞")
                        filtered_terms.append(term)
                else:
                    filtered_terms.append(term)
        
        return filtered_terms[:max_result_count] if max_result_count else filtered_terms
    
    def _should_filter_single_term(self, term: str) -> bool:
        """
        判斷單個詞是否應該過濾（用於 call_prompt 等單個字符串的過濾判斷）
        
        Args:
            term: 要判斷的詞或短語
            
        Returns:
            是否應該過濾（True=應該過濾，False=不應該過濾）
        """
        if not self.filter_casual_terms_enabled or not term:
            return False
        
        # 如果長度超過閾值，不需要過濾
        if len(term) > self.filter_casual_terms_threshold:
            return False
        
        # 使用 _filter_casual_terms 判斷
        filtered_result = self._filter_casual_terms([term], max_result_count=1)
        
        # 如果過濾後結果為空，表示應該過濾
        return len(filtered_result) == 0
    
    def _extract_system_collections(self, config: Dict[str, Any]) -> set:
        """
        從配置中提取系統儲存區（collection）名稱
        
        Args:
            config: 系統配置字典
            
        Returns:
            系統 collection 名稱集合（小寫，用於比對）
        """
        collections = set()
        
        # 從 api 配置中提取
        api_config = config.get('api', {})
        if 'default_collection' in api_config:
            collections.add(api_config['default_collection'].lower())
        if 'qa_default_collection' in api_config:
            collections.add(api_config['qa_default_collection'].lower())
        if 'add_point_default_collection' in api_config:
            collections.add(api_config['add_point_default_collection'].lower())
        
        question_analysis = api_config.get('question_analysis', {})
        if 'default_collection' in question_analysis:
            collections.add(question_analysis['default_collection'].lower())
        
        # 從 search_settings 配置中提取
        search_settings = config.get('search_settings', {})
        if 'collection_name' in search_settings:
            collections.add(search_settings['collection_name'].lower())
        
        # 從 crawler.sharepoint 配置中提取
        crawler = config.get('crawler', {})
        sharepoint = crawler.get('sharepoint', {})
        if 'default_collection' in sharepoint:
            collections.add(sharepoint['default_collection'].lower())
        
        # 從 collections.auto_create.patterns 中提取（過濾掉通配符模式）
        collections_config = config.get('collections', {})
        auto_create = collections_config.get('auto_create', {})
        patterns = auto_create.get('patterns', [])
        for pattern in patterns:
            # 只添加非通配符模式（不包含 * 或 ?）
            if isinstance(pattern, str) and '*' not in pattern and '?' not in pattern:
                collections.add(pattern.lower())
        
        # 添加常見的測試用 collection 名稱（避免污染文件內容）
        test_collections = {
            'test', 'test_update_flow', 'test_duplicate_path_flow',
            'test_batch', 'test_collection_basic', 'test_basic'
        }
        collections.update(test_collections)
        
        return collections
    
    def detect_pattern(self, text: str) -> Tuple[str, float]:
        """
        檢測文件書寫模式（LLM 輔助）
        
        Args:
            text: 文本內容
            
        Returns:
            (模式類型, 置信度) 元組
        """
        text_sample = text[:3000]

        # 先進行規則預檢
        rule_result = self._detect_with_rules(text_sample)
        
        # 如果規則檢測置信度高，直接返回
        if rule_result['confidence'] >= 0.8:
            return (rule_result['pattern'], rule_result['confidence'])
        
        # 否則使用 LLM 輔助檢測
        llm_result = self._detect_with_llm(text)
        
        # 合併結果（LLM 優先，必要時以規則結果做保險）
        if llm_result:
            if rule_result['pattern'] == self.PATTERN_PLAIN_TEXT:
                # 只有在 LLM 對非純文字模式非常有信心時才覆蓋
                if llm_result['pattern'] == self.PATTERN_STRUCTURED_DOC:
                    doc_score = self._detect_structured_doc_pattern(text_sample)
                    if doc_score < 0.3 and llm_result['confidence'] < 0.95:
                        return (rule_result['pattern'], rule_result['confidence'])
                return (llm_result['pattern'], llm_result['confidence'])
            return (llm_result['pattern'], llm_result['confidence'])
        
        # 如果 LLM 檢測失敗，使用規則檢測結果
        return (rule_result['pattern'], rule_result['confidence'])
    
    def _detect_with_rules(self, text: str) -> Dict[str, Any]:
        """
        規則檢測（快速預檢）
        
        Args:
            text: 文本內容
            
        Returns:
            包含 pattern 和 confidence 的字典
        """
        text_sample = text[:3000]  # 取前 3000 字符進行檢測
        
        # 檢測程式碼
        code_score = self._detect_code_pattern(text_sample)
        if code_score > 0.7:
            return {'pattern': self.PATTERN_CODE, 'confidence': code_score}
        
        # 檢測日誌
        log_score = self._detect_log_pattern(text_sample)
        if log_score > 0.7:
            return {'pattern': self.PATTERN_LOG, 'confidence': log_score}
        
        # 檢測溝通
        comm_score = self._detect_communication_pattern(text_sample)
        if comm_score > 0.7:
            return {'pattern': self.PATTERN_COMMUNICATION, 'confidence': comm_score}
        
        # 檢測結構化文檔
        doc_score = self._detect_structured_doc_pattern(text_sample)
        if doc_score > 0.7:
            return {'pattern': self.PATTERN_STRUCTURED_DOC, 'confidence': doc_score}
        
        # 預設為純文字
        return {'pattern': self.PATTERN_PLAIN_TEXT, 'confidence': 0.5}
    
    def _detect_code_pattern(self, text: str) -> float:
        """檢測程式碼模式"""
        score = 0.0
        
        # 檢查程式碼關鍵字
        code_keywords = ['def ', 'class ', 'function ', 'import ', 'return ', 
                        'const ', 'let ', 'var ', 'public ', 'private ']
        keyword_count = sum(1 for kw in code_keywords if kw in text)
        if keyword_count > 0:
            score += min(keyword_count * 0.2, 0.6)
        
        # 檢查程式碼結構（括號配對、縮排）
        if re.search(r'\{[^}]*\}|\[[^\]]*\]|\([^)]*\)', text):
            score += 0.2
        
        # 檢查註解符號
        if re.search(r'^\s*#|^\s*//|^\s*/\*', text, re.MULTILINE):
            score += 0.2
        
        return min(score, 1.0)
    
    def _detect_log_pattern(self, text: str) -> float:
        """檢測日誌模式"""
        score = 0.0
        
        # 檢查時間戳
        timestamp_patterns = [
            r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}',
            r'\[\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\]'
        ]
        if any(re.search(p, text) for p in timestamp_patterns):
            score += 0.4
        
        # 檢查日誌級別
        log_levels = ['INFO', 'ERROR', 'WARNING', 'WARN', 'DEBUG', 'TRACE', 'FATAL']
        if any(level in text for level in log_levels):
            score += 0.4
        
        # 檢查日誌格式特徵（流水帳式）
        lines = text.split('\n')
        if len(lines) > 5:
            # 檢查是否有重複的格式模式
            if len(set(lines[:10])) < len(lines[:10]) * 0.7:
                score += 0.2
        
        return min(score, 1.0)
    
    def _detect_communication_pattern(self, text: str) -> float:
        """檢測溝通模式"""
        score = 0.0
        
        # 檢查郵件頭
        email_headers = ['From:', 'To:', 'Subject:', 'Date:', 'Cc:', 'Bcc:']
        header_hits = sum(1 for header in email_headers if header.lower() in text.lower())
        if header_hits >= 2:
            score += 0.6
        elif header_hits == 1:
            score += 0.4

        # 檢查郵件地址
        if re.search(r'[\w\.-]+@[\w\.-]+', text):
            score += 0.3
        
        # 檢查聊天格式（發送者標識）
        if re.search(r'^\[.*\]:|^<.*>:|^.*\s+\d{2}:\d{2}', text, re.MULTILINE):
            score += 0.3
        
        # 檢查回覆關係
        if 'Re:' in text or 'RE:' in text or 'Fwd:' in text:
            score += 0.2
        
        return min(score, 1.0)
    
    def _detect_structured_doc_pattern(self, text: str) -> float:
        """檢測結構化文檔模式"""
        score = 0.0
        
        # 檢查 Markdown 標題
        if re.search(r'^#+\s+', text, re.MULTILINE):
            score += 0.4
        
        # 檢查列表
        if re.search(r'^\s*[-*+]\s+|^\s*\d+\.\s+', text, re.MULTILINE):
            score += 0.2
        
        # 檢查代碼塊
        if re.search(r'```|`', text):
            score += 0.2
        
        # 檢查段落結構（多個空行分隔）
        if re.search(r'\n\s*\n\s*\n', text):
            score += 0.2
        
        return min(score, 1.0)
    
    def _detect_with_llm(self, text: str) -> Optional[Dict[str, Any]]:
        """
        LLM 輔助檢測
        
        Args:
            text: 文本內容
            
        Returns:
            包含 pattern 和 confidence 的字典，失敗返回 None
        """
        if not self.llm_service or not self.pattern_config:
            return None
        
        try:
            # 準備提示詞
            system_prompt = self.pattern_config.get('system_prompt', '')
            user_prompt_template = self.pattern_config.get('user_prompt_template', '')
            max_length = self.pattern_config.get('max_content_length', 3000)
            
            # 截取文本
            text_sample = text[:max_length]
            
            # 構建用戶提示詞
            user_prompt = user_prompt_template.format(content=text_sample)
            
            # 調用 LLM（使用 TextProcessor 的 /chat API）
            # 使用 config.json 中的 LLM 配置
            response = self.llm_service.chat(
                prompt=user_prompt,
                system_prompt=system_prompt,
                provider=self.llm_provider,
                model=self.llm_model,
                max_tokens=self.llm_max_tokens,
                temperature=self.llm_temperature
            )
            
            if not response:
                self.logger.warning("LLM 回應為空")
                return None
            
            # 解析 JSON 回應
            # TextProcessor 的 /chat API 返回的 output 欄位可能包含：
            # 1. 純 JSON 字串
            # 2. Markdown 代碼塊包裹的 JSON
            # 3. 包含其他文字的 JSON
            
            result = None
            response_clean = response.strip()
            candidates = [response_clean]

            # 將常見的全形標點轉換為半形，增加解析成功率
            normalized = response_clean.replace('，', ',').replace('：', ':').replace('；', ';')
            if normalized != response_clean:
                candidates.append(normalized)
            
            for candidate in candidates:
                # 嘗試 1: 直接解析（如果是純 JSON）
                try:
                    result = json.loads(candidate)
                    break
                except json.JSONDecodeError:
                    pass
                
                # 嘗試 2: 提取 markdown 代碼塊中的 JSON
                code_block_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', candidate, re.DOTALL)
                if code_block_match:
                    try:
                        result = json.loads(code_block_match.group(1))
                        break
                    except json.JSONDecodeError:
                        pass
                
                # 嘗試 3: 提取第一個完整的 JSON 物件（使用更簡單的方法）
                start_idx = candidate.find('{')
                end_idx = candidate.rfind('}')
                if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
                    try:
                        json_str = candidate[start_idx:end_idx + 1]
                        result = json.loads(json_str)
                        break
                    except json.JSONDecodeError:
                        continue
            
            if result is None:
                self.logger.error(f"無法解析 LLM 回應為 JSON: {response[:200]}")
                return None
            
            # 驗證結果格式
            pattern = result.get('pattern', self.PATTERN_PLAIN_TEXT)
            confidence = float(result.get('confidence', 0.5))
            
            # 驗證模式類型
            valid_patterns = [
                self.PATTERN_CODE,
                self.PATTERN_LOG,
                self.PATTERN_COMMUNICATION,
                self.PATTERN_STRUCTURED_DOC,
                self.PATTERN_PLAIN_TEXT
            ]
            if pattern not in valid_patterns:
                pattern = self.PATTERN_PLAIN_TEXT
            
            return {
                'pattern': pattern,
                'confidence': confidence,
                'reasoning': result.get('reasoning', ''),
                'features': result.get('features', {})
            }
            
        except Exception as e:
            self.logger.error(f"LLM 模式檢測失敗: {e}")
            return None
    
    def segment(self, text: str, pattern: str = None, segments: Optional[List[Dict[str, Any]]]=None) -> bool:
        """
        根據模式進行分段
        
        Args:
            text: 文本內容
            pattern: 模式類型（如果為 None，則自動檢測）
            segments: 分段結果列表（可選，如果為 None 則創建新列表）
                     分段結果會填充到此列表中
            
        Returns:
            bool: 是否成功分段（True 表示成功，False 表示失敗）
                 分段結果通過 segments 參數返回
        """
        # 如果未指定模式，自動檢測
        if pattern is None:
            pattern, _ = self.detect_pattern(text)
        
        # 根據模式選擇分段策略
        if segments is None:
            segments = []
        if pattern == self.PATTERN_CODE:
            if not self._segment_code(text, segments):
                return False
        elif pattern == self.PATTERN_LOG:
            if not self._segment_log(text, segments):
                return False
        elif pattern == self.PATTERN_COMMUNICATION:
            if not self._segment_communication(text, segments):
                return False
        elif pattern == self.PATTERN_STRUCTURED_DOC:
            if not self._segment_structured_doc(text, segments):
                return False
        else:
            # 預設分段（按段落）
            if not self._segment_plain_text(text, segments):
                return False
        return True
    
    def _build_keyword_prompt(self, text: str) -> Dict[str, Any]:
        """
        构建關鍵詞提取的 prompt 字典（用於並行调用）
        
        Args:
            text: 文本内容
            
        Returns:
            包含 prompt 信息的字典
        """
        if not self.extract_keywords_config:
            return None
        
        text_clean = text.strip() if text else ''
        min_text_length = 10
        
        if len(text_clean) < min_text_length:
            return None
        
        system_prompt = self.extract_keywords_config.get('system_prompt', '')
        user_prompt_template = self.extract_keywords_config.get('user_prompt_template', '')
        text_sample = text_clean[:1000] if len(text_clean) > 1000 else text_clean
        text_sample = text_sample.replace('，', ',').replace('：', ':').replace('；', ';')
        user_prompt = user_prompt_template.format(content=text_sample)
        
        generation_config = self.extract_keywords_config.get('generation_config', {})
        max_tokens = int(generation_config.get('max_new_tokens', 50))
        temperature = float(generation_config.get('temperature', 0.3))
        
        return {
            'prompt': user_prompt,
            'system_prompt': system_prompt,
            'provider': self.llm_provider,
            'model': self.llm_model,
            'max_tokens': max_tokens,
            'temperature': temperature
        }

    def _normalize_multi_prompts(self, prompts: List[str]) -> List[str]:
        """Normalize multi_prompts: strip, drop empty, dedup exact (keep order)."""
        return _normalize_multi_prompts_list(prompts)

    def _build_dedup_multi_prompts_prompt(self, prompts: List[str]) -> Optional[Dict[str, Any]]:
        """Build LLM prompt for deduplicating similar multi_prompts (return indices)."""
        if not self.dedup_multi_prompts_config:
            return None
        if not prompts:
            return None

        system_prompt = self.dedup_multi_prompts_config.get('system_prompt', '')
        user_prompt_template = self.dedup_multi_prompts_config.get('user_prompt_template', '')
        items_json = json.dumps(prompts, ensure_ascii=False)
        user_prompt = user_prompt_template.format(items_json=items_json)

        generation_config = self.dedup_multi_prompts_config.get('generation_config', {})
        max_tokens = int(generation_config.get('max_new_tokens', 200))
        temperature = float(generation_config.get('temperature', 0.2))

        return {
            'prompt': user_prompt,
            'system_prompt': system_prompt,
            'provider': self.llm_provider,
            'model': self.llm_model,
            'max_tokens': max_tokens,
            'temperature': temperature
        }

    def _parse_dedup_multi_prompts_result(self, response: str, original: List[str]) -> List[str]:
        """解析 dedup_multi_prompts 回傳的內容 (success/reason/mask) 去選擇要保留的 multi prompt"""
        if not response:
            return original
        response_clean = response.strip()
        match = re.search(r'\{[\s\S]*\}', response_clean)
        if not match:
            return original
        try:
            parsed = json.loads(match.group(0))
        except Exception:
            return original
        if not isinstance(parsed, dict):
            return original

        success = parsed.get('success')
        reason = parsed.get('reason')
        if reason and self.logger:
            self.logger.info(f"[_dedup_multi_prompts_batch] LLM reason: {reason}")
        if success is not True:
            return original

        mask = parsed.get('mask')
        if not isinstance(mask, list) or len(mask) != len(original):
            return original
        for m in mask:
            if not isinstance(m, int) or m not in (0, 1):
                return original
        kept = [x for x, m in zip(original, mask) if m == 1]
        return kept if kept else original

    def _dedup_multi_prompts_batch(self, prompts_by_segment: List[List[str]]) -> List[List[str]]:
        """Deduplicate multi_prompts for many segments via batch LLM calls."""
        cleaned = [self._normalize_multi_prompts(p) for p in prompts_by_segment]
        if not self.dedup_multi_prompts_enabled:
            return cleaned
        if not self.parallel_executor or not self.llm_service:
            return cleaned

        prompts_list = []
        index_map = []
        for idx, items in enumerate(cleaned):
            if len(items) < self.dedup_multi_prompts_min_count:
                continue
            if self.dedup_multi_prompts_max_items > 0 and len(items) > self.dedup_multi_prompts_max_items:
                items = items[:self.dedup_multi_prompts_max_items]
            prompt_dict = self._build_dedup_multi_prompts_prompt(items)
            if not prompt_dict:
                continue
            prompts_list.append(prompt_dict)
            index_map.append((idx, items))

        if not prompts_list:
            return cleaned

        results = []
        batch_size = max(1, self.dedup_multi_prompts_batch_size)
        try:
            for start in range(0, len(prompts_list), batch_size):
                batch = prompts_list[start:start + batch_size]
                results.extend(self.parallel_executor.batch_chat(batch))
        except Exception as e:
            if self.logger:
                self.logger.warning(f"[_dedup_multi_prompts_batch] LLM 批量去重失敗: {e}")
            return cleaned

        for (idx, original_items), result in zip(index_map, results):
            if not result:
                continue
            deduped = self._parse_dedup_multi_prompts_result(result, original_items)
            if self.dedup_multi_prompts_max_keep > 0:
                deduped = deduped[:self.dedup_multi_prompts_max_keep]
            cleaned[idx] = deduped

        return cleaned

    def _dedup_segment_multi_prompts(self, segments: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Apply LLM-based dedup to segment multi_prompts in batch."""
        if not segments:
            return segments

        multi_prompts_list = []
        for seg in segments:
            mp = seg.get('multi_prompts')
            if isinstance(mp, dict):
                multi_prompts_list.append(list(mp.values()))
            elif isinstance(mp, list):
                multi_prompts_list.append(mp)
            else:
                multi_prompts_list.append([])

        deduped_list = self._dedup_multi_prompts_batch(multi_prompts_list)
        for seg, deduped in zip(segments, deduped_list):
            if deduped is not None:
                seg['multi_prompts'] = deduped
        return segments

    def _parse_keywords(self, response: str) -> List[str]:
        """
        解析 LLM 返回的關鍵詞（從响應字符串中提取）
        
        Args:
            response: LLM 响應字符串
            
        Returns:
            關鍵詞列表（最多5個），失败時返回空列表
        """
        if not response:
            return []
        
        response_clean = response.strip()
        
        # 檢測錯誤消息模式
        error_patterns = self.config.get('llm_error', {}).get('patterns', [])
        if any(pattern in response_clean for pattern in error_patterns):
            return []
        
        # 使用共用函數解析關鍵詞
        return parse_keywords_from_text(response_clean, max_keywords=5)
    
    def _build_filter_prompt(self, term: str) -> Dict[str, Any]:
        """
        構建過濾判斷的 prompt 字典（用於並行調用）
        
        Args:
            term: 要判斷的詞或短語
            
        Returns:
            包含 prompt 信息的字典
        """
        if not self.filter_casual_terms_config:
            return None
        
        term_clean = term.strip() if term else ''
        if not term_clean:
            return None
        
        system_prompt = self.filter_casual_terms_config.get('system_prompt', '')
        user_prompt_template = self.filter_casual_terms_config.get('user_prompt_template', '')
        user_prompt = user_prompt_template.format(term=term_clean)
        
        generation_config = self.filter_casual_terms_config.get('generation_config', {})
        max_tokens = int(generation_config.get('max_new_tokens', 50))
        temperature = float(generation_config.get('temperature', 0.2))
        
        return {
            'prompt': user_prompt,
            'system_prompt': system_prompt,
            'provider': self.llm_provider,
            'model': self.llm_model,
            'max_tokens': max_tokens,
            'temperature': temperature
        }
    
    def _parse_filter_result(self, response: str) -> bool:
        """
        解析 LLM 返回的過濾判斷結果
        
        Args:
            response: LLM 响應字符串
            
        Returns:
            是否應該過濾（True=應該過濾，False=不應該過濾），失敗時返回 False（保守策略）
        """
        if not response:
            return False
        
        response_clean = response.strip()
        
        # 檢測錯誤消息模式
        error_patterns = self.config.get('llm_error', {}).get('patterns', [])
        if any(pattern in response_clean for pattern in error_patterns):
            return False
        
        # 嘗試解析 JSON
        try:
            # 嘗試直接解析 JSON
            if response_clean.startswith('{') and response_clean.endswith('}'):
                result = json.loads(response_clean)
                if isinstance(result, dict):
                    should_filter = result.get('should_filter', False)
                    if result.get('reason'):    self.logger.info(f"[{result.get('should_filter')}]filter reason: {result.get('reason')}")
                    return bool(should_filter)
        except (json.JSONDecodeError, ValueError, TypeError):
            pass
        
        # 嘗試從 markdown 代碼塊中提取 JSON
        try:
            import re
            json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', response_clean, re.DOTALL)
            if json_match:
                result = json.loads(json_match.group(1))
                if isinstance(result, dict):
                    should_filter = result.get('should_filter', False)
                    if result.get('reason'):    self.logger.info(f"[{result.get('should_filter')}]filter reason: {result.get('reason')}")
                    return bool(should_filter)
        except (json.JSONDecodeError, ValueError, TypeError):
            pass
        
        # 嘗試提取第一個完整的 JSON 對象
        try:
            import re
            json_match = re.search(r'\{[^{}]*"should_filter"[^{}]*\}', response_clean)
            if json_match:
                result = json.loads(json_match.group(0))
                if isinstance(result, dict):
                    should_filter = result.get('should_filter', False)
                    if result.get('reason'):    self.logger.info(f"[{result.get('should_filter')}]filter reason: {result.get('reason')}")
                    return bool(should_filter)
        except (json.JSONDecodeError, ValueError, TypeError):
            pass
        
        # 如果無法解析，使用保守策略（不過濾）
        return False
    
    def _extract_content_keywords(self, text: str) -> List[str]:
        """
        使用 LLM 從文本內容中提取關鍵詞（內容標籤萃取）
        
        Args:
            text: 文本內容
            
        Returns:
            關鍵詞列表（最多5個），失敗時返回空列表
        """
        if not self.llm_service or not self.extract_keywords_config:
            self.logger.warning("LLM 服務或關鍵詞提取 prompt 配置未初始化")
            return []
        
        # 檢查文本是否足夠（避免 LLM 返回錯誤消息）
        text_clean = text.strip() if text else ''
        min_text_length = 10  # 最少 10 個字符
        
        if len(text_clean) < min_text_length:
            self.logger.debug(f"文本過短 ({len(text_clean)} 字符)，跳過關鍵詞提取")
            return []
        
        try:
            # 準備提示詞
            system_prompt = self.extract_keywords_config.get('system_prompt', '')
            user_prompt_template = self.extract_keywords_config.get('user_prompt_template', '')
            
            # 限制文本長度（前1000字符）避免過長
            text_sample = text_clean[:5000]
            text_sample = text_sample.replace('，', ',').replace('：', ':').replace('；', ';')
            
            # 構建用戶提示詞
            user_prompt = user_prompt_template.format(content=text_sample)
            
            # 獲取生成配置
            generation_config = self.extract_keywords_config.get('generation_config', {})
            max_tokens = int(generation_config.get('max_new_tokens', 50))
            temperature = float(generation_config.get('temperature', 0.3))
            
            # 調用 LLM
            response = self.llm_service.chat(
                prompt=user_prompt,
                system_prompt=system_prompt,
                provider=self.llm_provider,
                model=self.llm_model,
                max_tokens=max_tokens,
                temperature=temperature
            )
            
            if not response:
                _debug = {'system_prompt': system_prompt, 'user_prompt': user_prompt, 'response': response}
                m_debug.update(_debug)
                m_debug.save()
                self.logger.warning("LLM 關鍵詞提取回應為空")
                return []
            
            # 解析回應 - 使用共用函數清洗 LLM 回應
            response_clean = clean_llm_keyword_response(response)
            
            # 檢測錯誤消息模式（LLM 返回的錯誤提示）
            error_patterns = self.config.get('llm_error', {}).get('patterns', [])
            
            # 如果回應包含錯誤消息模式，直接返回空列表
            if any(pattern in response_clean for pattern in error_patterns):
                self.logger.warning(f"LLM 返回錯誤消息，跳過關鍵詞提取: {response_clean[:100]}...")
                return []
            
            keywords = []
            
            # 準備候選文本（原始和正規化版本）
            normalized = response_clean
            
            # 嘗試 1: 直接解析 JSON（如果回應是 JSON 格式）
            candidates = [response_clean, normalized]
            
            # 嘗試修復格式錯誤的 JSON（針對有額外空格的情況）
            for i, candidate in enumerate(candidates[:]):
                # 如果看起來像 JSON 陣列但可能有格式問題
                if candidate.strip().startswith('[') and candidate.strip().endswith(']'):
                    # 嘗試修復：清理所有引號周圍的空格
                    fixed = re.sub(r'\[\s*\"', '["', candidate)  # 開始處
                    fixed = re.sub(r'\"\s*\]', '"]', fixed)      # 結尾處
                    fixed = re.sub(r'\"\s*,\s*\"', '","', fixed) # 中間的逗號
                    if fixed != candidate:
                        candidates.append(fixed)
            
            for candidate in candidates:
                try:
                    keywords_data = json.loads(candidate)
                    # 處理 {"keywords": [...]} 格式
                    if isinstance(keywords_data, dict):
                        keywords = keywords_data.get('keywords', [])
                    # 處理直接返回列表 ["keyword1", "keyword2"] 格式
                    elif isinstance(keywords_data, list):
                        keywords = keywords_data
                    else:
                        keywords = []
                    
                    if isinstance(keywords, list) and keywords:
                        # 確保所有元素都是字符串，並過濾空值
                        keywords = [str(kw).strip() for kw in keywords if kw and len(str(kw).strip()) > 0]
                        keywords = keywords[:5]  # 限制最多5個關鍵詞
                        if keywords:
                            self.logger.debug(f"提取關鍵詞成功 (JSON): {', '.join(keywords)}")
                            return keywords
                except json.JSONDecodeError as je:
                    # 如果是 "Extra data" 錯誤，嘗試只解析第一個完整的 JSON 對象/數組
                    if "Extra data" in str(je):
                        try:
                            # 找到第一個完整的 JSON 數組或對象
                            json_start = candidate.find('[')
                            if json_start == -1:
                                json_start = candidate.find('{')
                            
                            if json_start != -1:
                                # 嘗試找到對應的結束符號
                                bracket_count = 0
                                json_end = -1
                                for i in range(json_start, len(candidate)):
                                    if candidate[i] == '[':
                                        bracket_count += 1
                                    elif candidate[i] == ']':
                                        bracket_count -= 1
                                        if bracket_count == 0:
                                            json_end = i
                                            break
                                
                                if json_end != -1:
                                    # 只解析第一個完整的 JSON
                                    json_part = candidate[json_start:json_end + 1]
                                    keywords_data = json.loads(json_part)
                                    if isinstance(keywords_data, dict):
                                        keywords = keywords_data.get('keywords', [])
                                    elif isinstance(keywords_data, list):
                                        keywords = keywords_data
                                    else:
                                        keywords = []
                                    
                                    if isinstance(keywords, list) and keywords:
                                        keywords = [str(kw).strip() for kw in keywords if kw and len(str(kw).strip()) > 0]
                                        keywords = keywords[:5]
                                        if keywords:
                                            self.logger.debug(f"提取關鍵詞成功 (JSON, 處理 Extra data): {', '.join(keywords)}")
                                            return keywords
                        except Exception:
                            pass
                    
                    # 記錄 JSON 解析錯誤的詳細信息（僅在最後一次嘗試時）
                    if candidate == candidates[-1]:
                        self.logger.debug(f"JSON 解析失敗: {je}, 內容: {candidate[:100]}...")
                    pass
            
            # 嘗試 2: 提取 markdown 代碼塊中的 JSON（支持 {} 和 [] 兩種格式）
            for candidate in [response_clean, normalized]:
                # 修改正則表達式以支持列表格式 [...] 和字典格式 {...}
                code_block_match = re.search(r'```(?:json)?\s*([{\[].*?[}\]])\s*```', candidate, re.DOTALL)
                if code_block_match:
                    try:
                        keywords_data = json.loads(code_block_match.group(1))
                        # 處理 {"keywords": [...]} 格式
                        if isinstance(keywords_data, dict):
                            keywords = keywords_data.get('keywords', [])
                        # 處理直接返回列表 ["keyword1", "keyword2"] 格式
                        elif isinstance(keywords_data, list):
                            keywords = keywords_data
                        else:
                            keywords = []
                        
                        if isinstance(keywords, list) and keywords:
                            keywords = [str(kw).strip() for kw in keywords if kw and len(str(kw).strip()) > 0]
                            keywords = keywords[:5]
                            if keywords:
                                self.logger.debug(f"提取關鍵詞成功 (Markdown JSON): {', '.join(keywords)}")
                                return keywords
                    except json.JSONDecodeError:
                        pass
            
            # 嘗試 3: 提取第一個完整的 JSON 物件或列表（支持 {} 和 [] 兩種格式）
            for candidate in [response_clean, normalized]:
                # 嘗試提取字典格式 {...}（使用配對括號查找）
                dict_start_idx = candidate.find('{')
                if dict_start_idx != -1:
                    # 找到配對的右括號
                    brace_count = 0
                    dict_end_idx = -1
                    for i in range(dict_start_idx, len(candidate)):
                        if candidate[i] == '{':
                            brace_count += 1
                        elif candidate[i] == '}':
                            brace_count -= 1
                            if brace_count == 0:
                                dict_end_idx = i
                                break
                    
                    if dict_end_idx != -1 and dict_end_idx > dict_start_idx:
                        try:
                            json_str = candidate[dict_start_idx:dict_end_idx + 1]
                            keywords_data = json.loads(json_str)
                            if isinstance(keywords_data, dict):
                                keywords = keywords_data.get('keywords', [])
                                if isinstance(keywords, list) and keywords:
                                    keywords = [str(kw).strip() for kw in keywords if kw and len(str(kw).strip()) > 0]
                                    keywords = keywords[:5]
                                    if keywords:
                                        self.logger.debug(f"提取關鍵詞成功 (提取 JSON dict): {', '.join(keywords)}")
                                        return keywords
                        except json.JSONDecodeError:
                            pass
                
                # 嘗試提取列表格式 [...]（使用配對括號查找）
                list_start_idx = candidate.find('[')
                if list_start_idx != -1:
                    # 找到配對的右括號
                    bracket_count = 0
                    list_end_idx = -1
                    for i in range(list_start_idx, len(candidate)):
                        if candidate[i] == '[':
                            bracket_count += 1
                        elif candidate[i] == ']':
                            bracket_count -= 1
                            if bracket_count == 0:
                                list_end_idx = i
                                break
                    
                    if list_end_idx != -1 and list_end_idx > list_start_idx:
                        try:
                            json_str = candidate[list_start_idx:list_end_idx + 1]
                            keywords_data = json.loads(json_str)
                            if isinstance(keywords_data, list) and keywords_data:
                                keywords = [str(kw).strip() for kw in keywords_data if kw and len(str(kw).strip()) > 0]
                                keywords = keywords[:5]
                                if keywords:
                                    self.logger.debug(f"提取關鍵詞成功 (提取 JSON list): {', '.join(keywords)}")
                                    return keywords
                        except json.JSONDecodeError:
                            pass
            
            # 嘗試 4: 從文本中提取關鍵詞（支持多種分隔符）
            # 但如果文本看起來像 JSON 格式（以 [ 或 { 開頭），跳過分隔符分割
            # 因為這可能導致提取到帶引號的片段（如 '["URL"'）
            if not (response_clean.strip().startswith('[') or response_clean.strip().startswith('{')):
                # 先處理多字符分隔符（如 ||）
                if '||' in response_clean:
                    parts = response_clean.split('||')
                    keywords = [kw.strip() for kw in parts if kw.strip()]
                    # 過濾掉太短或太長的關鍵詞（可能是解析錯誤）
                    # 同時過濾掉看起來像 JSON 片段的部分（包含引號或中括號）
                    keywords = [kw for kw in keywords 
                              if 1 <= len(kw) <= 50 
                              and not (kw.startswith('["') or kw.startswith('[\'') or kw.startswith('{'))]
                    keywords = keywords[:5]  # 限制最多5個關鍵詞
                    if keywords:
                        self.logger.debug(f"提取關鍵詞成功 (分隔符 '||'): {', '.join(keywords)}")
                        return keywords
                
                # 支持的分隔符：逗號、中文頓號、分號、換行、空格（多個）
                separators = [
                    ',',      # 英文逗號
                    '，',     # 中文逗號
                    '、',     # 中文頓號
                    ';',      # 英文分號
                    '；',     # 中文分號
                    '\n',     # 換行
                    '\r\n',   # Windows 換行
                ]
                
                for sep in separators:
                    if sep in response_clean:
                        # 使用正則表達式分割，支持多個連續分隔符
                        parts = re.split(f'[{re.escape(sep)}\\s]+', response_clean)
                        keywords = [kw.strip() for kw in parts if kw.strip()]
                        # 過濾掉太短或太長的關鍵詞（可能是解析錯誤）
                        # 同時過濾掉看起來像 JSON 片段的部分（包含引號或中括號）
                        keywords = [kw for kw in keywords 
                                  if 1 <= len(kw) <= 50 
                                  and not (kw.startswith('["') or kw.startswith('[\'') or kw.startswith('{'))]
                        keywords = keywords[:5]  # 限制最多5個關鍵詞
                        if keywords:
                            self.logger.debug(f"提取關鍵詞成功 (分隔符 '{sep}'): {', '.join(keywords)}")
                            return keywords
            
            # 嘗試 5: 從列表格式中提取（例如：1. keyword1 2. keyword2 或 - keyword1 - keyword2）
            list_patterns = [
                r'[0-9一二三四五六七八九十]+[\.、]\s*([^\n]+)',  # 數字列表
                r'[-*+]\s*([^\n]+)',  # 符號列表
                r'•\s*([^\n]+)',       # 項目符號
            ]
            
            for pattern in list_patterns:
                matches = re.findall(pattern, response_clean)
                if matches:
                    keywords = [kw.strip() for kw in matches if kw.strip()]
                    keywords = [kw for kw in keywords if 1 <= len(kw) <= 50]
                    keywords = keywords[:5]
                    if keywords:
                        self.logger.debug(f"提取關鍵詞成功 (列表格式): {', '.join(keywords)}")
                        return keywords
            
            # 嘗試 6: 如果回應看起來像是一個簡單的關鍵詞（沒有分隔符）
            if len(response_clean) <= 50 and not any(sep in response_clean for sep in [',', '，', '、', ';', '；', '\n']):
                keywords = [response_clean.strip()]
                if keywords[0]:
                    self.logger.debug(f"提取關鍵詞成功 (單一關鍵詞): {keywords[0]}")
                    return keywords
            
            # 如果都失敗，記錄警告但嘗試提取前幾個詞作為關鍵詞
            # 使用空格或標點符號分割，取前5個非空詞
            fallback_keywords = re.split(r'[\s,，、;；\n]+', response_clean)
            fallback_keywords = [kw.strip() for kw in fallback_keywords if kw.strip() and 1 <= len(kw.strip()) <= 50]
            
            # 過濾掉說明性文字（包含特定關鍵字的片段）
            explanation_patterns = [
                '以下',
                '關鍵詞',
                '關鍵片語',
                '提取',
                '段落',
                '代表',
                '主旨',
                '準確',
                '列表',
                '如下'
            ]
            
            filtered_keywords = []
            for kw in fallback_keywords:
                # 如果關鍵詞包含說明性詞語（且長度超過10個字），跳過
                is_explanation = False
                if len(kw) > 10:
                    for pattern in explanation_patterns:
                        if pattern in kw:
                            is_explanation = True
                            break
                
                if not is_explanation:
                    filtered_keywords.append(kw)
            
            fallback_keywords = filtered_keywords[:5]
            
            if fallback_keywords:
                self.logger.warning(f"無法完全解析關鍵詞回應，使用備用方法提取: {response_clean[:200]}")
                self.logger.debug(f"備用提取結果: {', '.join(fallback_keywords)}")
                return fallback_keywords
            else:
                self.logger.warning(f"無法解析關鍵詞回應: {response_clean[:200]}")
                return []
            
        except Exception as e:
            self.logger.error(f"LLM 關鍵詞提取失敗: {e}")
            return []
    
    def _filter_meaningless_tags(self, tags: List[str]) -> List[str]:
        """
        過濾掉無意義的標籤（系統常用的檔案結構名稱、數字流水號等）
        委託模組層 filter_meaningless_tags 共用核心邏輯。
        """
        return filter_meaningless_tags(tags, self.config)
    
    def _extract_contextual_tags_by_backscan(
            self, 
            segments: List[Dict[str, Any]], 
            current_idx: int, 
            level_key: str = 'level',
            heading_key: str = 'heading'
        ) -> List[str]:
        """
        依規則為 segments[current_idx] 生成脈絡標籤（結構型標籤）：
        - 對於每個目標層級（從 current_level-1 到 0），往前找最接近的父節點
        - 優先選擇 level == target_level 的節點（精確匹配）
        - 如果找不到精確匹配，選擇最接近的較淺層級節點
        - 確保不會跳過中間層級
        回傳順序：由近到遠（父→祖父→曾祖父...）
        
        Args:
            segments: 分段結果列表
            current_idx: 當前 segment 的索引
            level_key: metadata 中層級字段的名稱（預設 'level'）
            heading_key: metadata 中標題字段的名稱（預設 'heading'）
            
        Returns:
            脈絡標籤列表（由近到遠）
        """
        if current_idx < 0 or current_idx >= len(segments):
            return []
        
        current_segment = segments[current_idx]
        current_metadata = current_segment.get('metadata', {})
        current_level = current_metadata.get(level_key)
        
        # 如果當前 segment 沒有層級信息，無法回溯
        if current_level is None:
            return []
        
        try:
            current_level = int(current_level)
        except (ValueError, TypeError):
            return []
        
        # 改進的回溯邏輯：對於每個目標層級（從 current_level-1 到 0），往前找最接近的父節點
        tags = []
        
        # 從 current_level - 1 開始，逐層往上找父節點
        for target_level in range(current_level - 1, -1, -1):
            # 往前回溯，找到第一個 level <= target_level 的節點
            # 優先選擇 level == target_level 的節點（精確匹配）
            best_match = None
            best_match_level = None
            j = current_idx - 1
            
            while j >= 0:
                prev_segment = segments[j]
                prev_metadata = prev_segment.get('metadata', {})
                prev_level = prev_metadata.get(level_key)
                
                if prev_level is not None:
                    try:
                        prev_level = int(prev_level)
                        # 如果找到精確匹配的層級，直接使用
                        if prev_level == target_level:
                            heading = prev_metadata.get(heading_key)
                            if heading and isinstance(heading, str) and heading.strip():
                                tags.append(heading.strip())
                                break  # 找到精確匹配，停止搜索這個目標層級
                        # 如果還沒找到精確匹配，記錄最接近的候選（level < target_level 且最大）
                        elif prev_level < target_level:
                            if best_match is None or prev_level > best_match_level:
                                best_match = prev_metadata.get(heading_key)
                                best_match_level = prev_level
                    except (ValueError, TypeError):
                        pass
                
                j -= 1
            
            # 如果沒有找到精確匹配，但有候選節點，使用候選節點
            if best_match and isinstance(best_match, str) and best_match.strip():
                # 只有在還沒添加過這個標題時才添加（避免重複）
                if best_match.strip() not in tags:
                    tags.append(best_match.strip())
        
        return tags
    
    def _detect_log_format_with_llm(self, text: str) -> Dict[str, Any]:
        """
        使用 LLM 偵測日誌格式（時間戳格式、日誌級別、分解符）
        
        Args:
            text: 日誌文本內容
            
        Returns:
            包含 timestamp_formats, log_levels, delimiters 的字典
        """
        if not self.llm_service or not self.log_format_analysis_config:
            self.logger.warning("LLM 服務或 prompt 配置未初始化，使用預設值")
            return {
                'timestamp_formats': [],
                'log_levels': [],
                'delimiters': [],
                'confidence': 0.0
            }
        
        try:
            # 準備提示詞
            system_prompt = self.log_format_analysis_config.get('system_prompt', '')
            user_prompt_template = self.log_format_analysis_config.get('user_prompt_template', '')
            max_length = self.log_format_analysis_config.get('max_content_length', 3000)
            
            # 截取文本
            text_sample = text[:max_length]
            
            # 構建用戶提示詞
            user_prompt = user_prompt_template.format(content=text_sample)
            
            # 調用 LLM
            response = self.llm_service.chat(
                prompt=user_prompt,
                system_prompt=system_prompt,
                provider=self.llm_provider,
                model=self.llm_model,
                max_tokens=self.llm_max_tokens,
                temperature=self.llm_temperature
            )
            
            if not response:
                self.logger.warning("LLM 回應為空，使用預設值")
                return {
                    'timestamp_formats': [],
                    'log_levels': [],
                    'delimiters': [],
                    'confidence': 0.0
                }
            
            # 解析 JSON 回應
            result = None
            response_clean = response.strip()
            
            # 嘗試 1: 直接解析（如果是純 JSON）
            try:
                result = json.loads(response_clean)
            except json.JSONDecodeError:
                pass
            
            # 嘗試 2: 提取 markdown 代碼塊中的 JSON
            if result is None:
                code_block_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', response_clean, re.DOTALL)
                if code_block_match:
                    try:
                        result = json.loads(code_block_match.group(1))
                    except json.JSONDecodeError:
                        pass
            
            # 嘗試 3: 提取第一個完整的 JSON 物件
            if result is None:
                start_idx = response_clean.find('{')
                end_idx = response_clean.rfind('}')
                if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
                    try:
                        json_str = response_clean[start_idx:end_idx + 1]
                        result = json.loads(json_str)
                    except json.JSONDecodeError:
                        pass
            
            if result is None:
                self.logger.error(f"無法解析 LLM 回應為 JSON: {response[:200]}")
                return {
                    'timestamp_formats': [],
                    'log_levels': [],
                    'delimiters': [],
                    'confidence': 0.0
                }
            
            # 驗證並返回結果
            return {
                'timestamp_formats': result.get('timestamp_formats', []),
                'log_levels': result.get('log_levels', []),
                'delimiters': result.get('delimiters', []),
                'confidence': float(result.get('confidence', 0.0))
            }
            
        except Exception as e:
            self.logger.error(f"LLM 日誌格式偵測失敗: {e}")
            return {
                'timestamp_formats': [],
                'log_levels': [],
                'delimiters': [],
                'confidence': 0.0
            }
    
    def _segment_log(self, text: str, segments: Optional[List[Dict[str, Any]]]=None) -> bool:
        """
        日誌分段
        
        Args:
            text: 文本內容
            segments: 分段結果列表（可選，如果為 None 則創建新列表）
            
        Returns:
            是否成功
        """
        if segments is None:
            segments = []
        config = self.log_config.get('segmentation_rules', {})
        min_length = config.get('min_segment_length', 10)
        max_length = config.get('max_segment_length', 3000)
        lines = text.split('\n')
        
        # 使用 LLM 偵測時間戳格式、日誌級別和分解符（已關閉）
        # format_analysis = self._detect_log_format_with_llm(text)
        # timestamp_formats = format_analysis.get('timestamp_formats', [])
        # log_levels = format_analysis.get('log_levels', [])
        # delimiters = format_analysis.get('delimiters', [])
        # confidence = format_analysis.get('confidence', 0.0)
        
        # self.logger.info(f"日誌格式分析結果: 時間戳格式={len(timestamp_formats)}個, 日誌級別={len(log_levels)}個, 分解符={len(delimiters)}個, 信心值={confidence:.2f}")
        
        # 直接使用預設值（跳過 LLM 偵測）
        timestamp_formats = [
            r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}',
            r'\[\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\]'
        ]
        log_levels = ['INFO', 'ERROR', 'WARNING', 'WARN', 'DEBUG', 'TRACE', 'FATAL', 'CRITICAL']
        delimiters = []
        
        # 按時間戳分段（優先級最高）
        if config.get('by_timestamp', {}).get('enabled', True):
            current_segment = []
            current_timestamp = None
            rule_config = config.get('by_timestamp', {})
            rule_max_length = rule_config.get('max_length', max_length)
            rule_min_length = rule_config.get('min_length', min_length)
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # 檢測時間戳
                timestamp = None
                for pattern in timestamp_formats:
                    match = re.search(pattern, line)
                    if match:
                        timestamp = match.group(0)
                        break
                
                if timestamp:
                    # 遇到新時間戳，保存前一段
                    if current_segment and current_timestamp:
                        content = '\n'.join(current_segment)
                        if len(content) >= rule_min_length:
                            segments.append({
                                'content': content,
                                'metadata': {
                                    'type': 'log_entry',
                                    'timestamp': current_timestamp,
                                    'timestamp_formats': timestamp_formats,
                                    'log_levels': log_levels,
                                    'delimiters': delimiters
                                }
                            })
                    current_segment = [line]
                    current_timestamp = timestamp
                else:
                    # 檢查長度限制
                    potential_content = '\n'.join(current_segment + [line])
                    if len(potential_content) > rule_max_length and current_segment:
                        content = '\n'.join(current_segment)
                        if len(content) >= rule_min_length:
                            segments.append({
                                'content': content,
                                'metadata': {
                                    'type': 'log_entry',
                                    'timestamp': current_timestamp,
                                    'timestamp_formats': timestamp_formats,
                                    'log_levels': log_levels,
                                    'delimiters': delimiters
                                }
                            })
                        current_segment = [line]
                    else:
                        current_segment.append(line)
            
            # 處理最後一段
            if current_segment:
                content = '\n'.join(current_segment)
                if len(content) >= rule_min_length:
                        segments.append({
                            'content': content,
                            'metadata': {
                                'type': 'log_entry',
                                'timestamp': current_timestamp,
                                'timestamp_formats': timestamp_formats,
                                'log_levels': log_levels,
                                'delimiters': delimiters
                            }
                        })
        
        # 按日誌級別分段
        if not segments and config.get('by_log_level', {}).get('enabled', False):
            current_segment = []
            current_level = None
            rule_config = config.get('by_log_level', {})
            rule_max_length = rule_config.get('max_length', max_length)
            rule_min_length = rule_config.get('min_length', min_length)
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # 檢測日誌級別
                level = None
                for log_level in log_levels:
                    if log_level in line:
                        level = log_level
                        break
                
                if level and level != current_level:
                    # 遇到新級別，保存前一段
                    if current_segment and current_level:
                        content = '\n'.join(current_segment)
                        if len(content) >= rule_min_length:
                            segments.append({
                                'content': content,
                                'metadata': {
                                    'type': 'log_entry',
                                    'log_level': current_level,
                                    'timestamp_formats': timestamp_formats,
                                    'log_levels': log_levels,
                                    'delimiters': delimiters
                                }
                            })
                    current_segment = [line]
                    current_level = level
                else:
                    # 檢查長度限制
                    potential_content = '\n'.join(current_segment + [line])
                    if len(potential_content) > rule_max_length and current_segment:
                        content = '\n'.join(current_segment)
                        if len(content) >= rule_min_length:
                            segments.append({
                                'content': content,
                                'metadata': {
                                    'type': 'log_entry',
                                    'log_level': current_level,
                                    'timestamp_formats': timestamp_formats,
                                    'log_levels': log_levels,
                                    'delimiters': delimiters
                                }
                            })
                        current_segment = [line]
                    else:
                        current_segment.append(line)
            
            # 處理最後一段
            if current_segment:
                content = '\n'.join(current_segment)
                if len(content) >= rule_min_length:
                    segments.append({
                        'content': content,
                        'metadata': {
                            'type': 'log_entry',
                            'log_level': current_level
                        }
                    })
        
        # 按行分段（每行日誌為一段）
        if not segments and config.get('by_line', {}).get('enabled', False):
            rule_config = config.get('by_line', {})
            rule_min_length = rule_config.get('min_length', min_length)
            rule_max_length = rule_config.get('max_length', max_length)
            
            for line in lines:
                line = line.strip()
                if line and len(line) >= rule_min_length and len(line) <= rule_max_length:
                    segments.append({
                        'content': line,
                        'metadata': {
                            'type': 'log_line',
                            'timestamp_formats': timestamp_formats,
                            'log_levels': log_levels,
                            'delimiters': delimiters
                        }
                    })
        
        # 按事件分段（相關日誌行歸為一段）
        if not segments and config.get('by_event', {}).get('enabled', True):
            current_segment = []
            rule_config = config.get('by_event', {})
            rule_max_length = rule_config.get('max_length', max_length)
            rule_min_length = rule_config.get('min_length', min_length)
            
            for line in lines:
                line = line.strip()
                if not line:
                    # 空行可能表示事件結束
                    if current_segment:
                        content = '\n'.join(current_segment)
                        if len(content) >= rule_min_length:
                            segments.append({
                                'content': content,
                                'metadata': {
                                    'type': 'log_event'
                                }
                            })
                        current_segment = []
                    continue
                
                # 檢查長度限制
                potential_content = '\n'.join(current_segment + [line])
                if len(potential_content) > rule_max_length and current_segment:
                    content = '\n'.join(current_segment)
                    if len(content) >= rule_min_length:
                        segments.append({
                            'content': content,
                            'metadata': {
                                'type': 'log_event',
                                'timestamp_formats': timestamp_formats,
                                'log_levels': log_levels,
                                'delimiters': delimiters
                            }
                        })
                    current_segment = [line]
                else:
                    current_segment.append(line)
            
            # 處理最後一段
            if current_segment:
                content = '\n'.join(current_segment)
                if len(content) >= rule_min_length:
                    segments.append({
                        'content': content,
                        'metadata': {
                            'type': 'log_event'
                        }
                    })
        
        # 如果沒有分段，按段落分段
        if not segments:
            self._segment_plain_text(text, segments)
        
        return True
    
    def _segment_code(self, text: str, segments: Optional[List[Dict[str, Any]]]=None) -> bool:
        """程式碼分段"""
        if segments is None:
            segments = []
        config = self.code_config.get('segmentation_rules', {})
        min_length = config.get('min_segment_length', 50)
        max_length = config.get('max_segment_length', 2000)
        
        # 按函數/方法分段
        if config.get('function_level', {}).get('enabled', True):
            # 匹配函數定義（Python: def, JavaScript: function, Java: public/private method）
            # 使用更精確的模式，匹配到函數定義行
            function_pattern = r'^\s*(def|function|async\s+def|public|private|protected)\s+\w+'
            matches = list(re.finditer(function_pattern, text, re.MULTILINE))
            
            if matches:
                for i, match in enumerate(matches):
                    start = match.start()
                    # 找到下一個函數定義的位置，或文件結尾
                    end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
                    content = text[start:end].strip()
                    
                    if content and len(content) >= min_length:
                        # 提取函數名作為 heading
                        function_name_match = re.search(r'(?:def|function|async\s+def|public|private|protected)\s+(\w+)', content)
                        function_name = function_name_match.group(1) if function_name_match else None
                        segments.append({
                            'content': content,
                            'metadata': {
                                'type': 'function',
                                'level': 1,  # 函數層級為 1
                                'heading': function_name,
                                'start': start,
                                'end': end
                            }
                        })
        
        # 如果沒有找到函數，按類別分段
        if not segments and config.get('class_level', {}).get('enabled', True):
            class_pattern = r'^\s*class\s+\w+'
            matches = list(re.finditer(class_pattern, text, re.MULTILINE))
            
            if matches:
                for i, match in enumerate(matches):
                    start = match.start()
                    end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
                    content = text[start:end].strip()
                    
                    if content and len(content) >= min_length:
                        # 提取類名作為 heading
                        class_name_match = re.search(r'class\s+(\w+)', content)
                        class_name = class_name_match.group(1) if class_name_match else None
                        segments.append({
                            'content': content,
                            'metadata': {
                                'type': 'class',
                                'level': 0,  # 類層級為 0（最外層）
                                'heading': class_name,
                                'start': start,
                                'end': end
                            }
                        })
        
        # 若以上都沒有匹配，按段落分段
        if not segments:
            self._segment_plain_text(text, segments)
        
        return True
    
    def _segment_communication(self, text: str, segments: Optional[List[Dict[str, Any]]]=None) -> bool:
        """溝通分段"""
        if segments is None:
            segments = []
        config = self.communication_config.get('segmentation_rules', {})
        lines = text.split('\n')
        
        # 按發送者分段
        if config.get('by_sender', {}).get('enabled', True):
            current_segment = []
            current_sender = None
            max_length = config.get('max_segment_length', 2000)
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # 檢測發送者（郵件頭或聊天格式）
                sender_match = re.search(
                    r'^(From:|To:|\<.*\>|\[.*\]:|.*\s+\d{2}:\d{2})',
                    line
                )
                
                if sender_match:
                    sender = sender_match.group(0)
                    if sender != current_sender:
                        if current_segment:
                            content = '\n'.join(current_segment)
                            if len(content) >= config.get('min_segment_length', 10):
                                segments.append({
                                    'content': content,
                                    'metadata': {
                                        'type': 'message',
                                        'sender': current_sender
                                    }
                                })
                        current_segment = [line]
                        current_sender = sender
                    else:
                        current_segment.append(line)
                else:
                    # 檢查長度限制
                    potential_content = '\n'.join(current_segment + [line])
                    if len(potential_content) > max_length and current_segment:
                        content = '\n'.join(current_segment)
                        if len(content) >= config.get('min_segment_length', 10):
                            segments.append({
                                'content': content,
                                'metadata': {
                                    'type': 'message',
                                    'sender': current_sender
                                }
                            })
                        current_segment = [line]
                    else:
                        current_segment.append(line)
            
            # 處理最後一段
            if current_segment:
                content = '\n'.join(current_segment)
                if len(content) >= config.get('min_segment_length', 10):
                    segments.append({
                        'content': content,
                        'metadata': {
                            'type': 'message',
                            'sender': current_sender
                        }
                    })
        
        # 如果沒有分段，按段落分段
        if not segments:
            self._segment_plain_text(text, segments)
        
        return True
    
    def _segment_structured_doc(self, text: str, segments: Optional[List[Dict[str, Any]]]=None) -> bool:
        """結構化文檔分段"""
        if segments is None:
            segments = []
        config = self.structured_doc_config.get('segmentation_rules', {})
        lines = text.split('\n')
        
        # 獲取標題模式配置（優先使用配置文件，否則使用預設值）
        heading_patterns_config = self.structured_doc_config.get('heading_patterns', {})
        markdown_patterns = heading_patterns_config.get('markdown', [])
        
        # 編譯標題正則表達式
        # ATX 標題模式：^#+\s+
        atx_pattern = None
        for pattern in markdown_patterns:
            if pattern.startswith('^#+'):
                atx_pattern = re.compile(pattern + r'(.+)$')
                break
        if atx_pattern is None:
            # 如果配置中沒有，使用預設的 ATX 模式
            atx_pattern = re.compile(r'^(#+)\s+(.+)$')
        
        # Setext 標題模式：^={3,}$ 或 ^-{3,}$
        setext_equals_pattern = re.compile(r'^={3,}$')
        setext_dashes_pattern = re.compile(r'^-{3,}$')
        
        # 按標題分段
        if config.get('by_heading', {}).get('enabled', True):
            current_segment = []
            current_heading = None
            current_heading_level = None
            max_length = config.get('max_segment_length', 3000)
            
            i = 0
            while i < len(lines):
                line = lines[i]
                is_heading = False
                heading_text = None
                heading_level = None
                is_setext = False
                
                # 檢測 Setext 標題（先檢測，因為需要檢查上一行）
                # Setext 格式：標題文字 + 下一行 === 或 ---
                if i > 0:
                    prev_line_stripped = lines[i - 1].strip()
                    line_stripped = line.strip()
                    
                    # 檢查當前行是否為 Setext 標題下劃線
                    if setext_equals_pattern.match(line_stripped):
                        # === 表示一級標題
                        if prev_line_stripped and not prev_line_stripped.startswith('#'):
                            is_heading = True
                            is_setext = True
                            heading_text = prev_line_stripped
                            heading_level = 1
                    elif setext_dashes_pattern.match(line_stripped):
                        # --- 表示二級標題
                        if prev_line_stripped and not prev_line_stripped.startswith('#'):
                            is_heading = True
                            is_setext = True
                            heading_text = prev_line_stripped
                            heading_level = 2
                
                # 檢測 ATX 標題（# 標題）
                if not is_heading:
                    atx_match = atx_pattern.match(line)
                    if atx_match:
                        is_heading = True
                        # 提取標題文字和層級
                        if len(atx_match.groups()) >= 2:
                            # 格式：^(#+)\s+(.+)$
                            heading_level = len(atx_match.group(1))
                            heading_text = atx_match.group(2)
                        else:
                            # 格式：^#+\s+(.+)$（只有一個 group）
                            heading_text = atx_match.group(1) if atx_match.groups() else line.lstrip('#').strip()
                            heading_level = len(line) - len(line.lstrip('#'))
                
                if is_heading:
                    # 遇到新標題，保存前一段
                    if current_segment:
                        # 如果是 Setext 標題，需要確保上一行（標題文字）不在 current_segment 中
                        if is_setext and current_segment and current_segment[-1] == lines[i - 1]:
                            current_segment.pop()
                        
                        content = '\n'.join(current_segment)
                        if len(content) >= config.get('min_segment_length', 50):
                            segments.append({
                                'content': content,
                                'metadata': {
                                    'type': 'section',
                                    'heading': current_heading,
                                    'level': current_heading_level
                                }
                            })
                    
                    # 開始新段落
                    if is_setext:
                        # Setext 標題：包含上一行（標題文字）和當前行（下劃線）
                        current_segment = [lines[i - 1], line]
                    else:
                        # ATX 標題：只包含當前行
                        current_segment = [line]
                    
                    current_heading = heading_text
                    current_heading_level = heading_level
                else:
                    # 檢查長度限制
                    potential_content = '\n'.join(current_segment + [line])
                    if len(potential_content) > max_length and current_segment:
                        content = '\n'.join(current_segment)
                        if len(content) >= config.get('min_segment_length', 50):
                            segments.append({
                                'content': content,
                                'metadata': {
                                    'type': 'section',
                                    'heading': current_heading,
                                    'level': current_heading_level
                                }
                            })
                        current_segment = [line]
                    else:
                        current_segment.append(line)
                
                i += 1
            
            # 處理最後一段
            if current_segment:
                content = '\n'.join(current_segment)
                if len(content) >= config.get('min_segment_length', 50):
                    # 計算當前標題的層級（如果有的話）
                    prev_level = current_heading_level if 'current_heading_level' in locals() else None
                    segments.append({
                        'content': content,
                        'metadata': {
                            'type': 'section',
                            'heading': current_heading,
                            'level': prev_level
                        }
                    })
        
        # fallback when heading segmentation fails
        if not segments:
            if not self._segment_plain_text(text, segments):
                return False
        
        return True
    
    def _segment_plain_text(self, text: str, segments: Optional[List[Dict[str, Any]]]=None) -> bool:
        """純文字分段（預設）"""
        if segments is None:
            segments = []
        paragraphs = text.split('\n\n')
        
        for para in paragraphs:
            para = para.strip()
            if para:
                segments.append({
                    'content': para,
                    'metadata': {'type': 'paragraph'}
                })
        
        return True
    
    def _extract_os_tags_from_path(
            self,
            os_tags: List[str],
            file_path: str,
            path_sep: str,
            max_content_tags: int,
        ) -> bool:
        """
        從文件路徑提取 OS 標籤並添加到 os_tags 列表中
        
        Args:
            os_tags: 檔案路徑相關標籤列表（會被修改）
            file_path: 文件路徑
            path_sep: 路徑分隔符
            max_content_tags: 最大內容標籤數量
        """
        abs_file_path = os.path.abspath(file_path)
        # 保存原始文件名（保持大小写）
        original_filename = os.path.basename(abs_file_path)
        path_parts = abs_file_path.split(path_sep)
        # 過濾掉空的部分
        path_parts = [part for part in path_parts if part]
        
        # 如果路徑在系統儲存路徑下，移除系統儲存路徑前綴部分
        if _is_system_storage_path(abs_file_path):
            # 找出系統儲存路徑的前綴部分，只保留相對路徑部分
            # 使用 normcase 進行路徑匹配（不區分大小寫），但保留原始路徑用於提取文件名
            normalized_path = os.path.normcase(abs_file_path).rstrip(path_sep)
            for root in m_storage_roots:
                if normalized_path.startswith(root + path_sep):
                    # 從原始路徑（非規範化）中提取相對路徑，保持原始大小寫
                    root_len = len(root + path_sep)
                    relative_path = abs_file_path[root_len:] if len(abs_file_path) > root_len else ''
                    if relative_path:
                        # 重新分割相對路徑（保持原始大小寫）
                        path_parts = relative_path.split(path_sep)
                        path_parts = [part for part in path_parts if part]
                    else:
                        # 如果相對路徑為空，使用原始檔案名稱（保持大小寫）
                        path_parts = [original_filename] if original_filename else []
                    break
        
        # 過濾掉系統目錄名稱（使用 _filter_meaningless_tags 來過濾）
        # 但保留檔案名稱（保持原始大小寫）
        filename = original_filename
        filtered_path_parts = self._filter_meaningless_tags(path_parts) if hasattr(self, '_filter_meaningless_tags') else path_parts
        
        # 如果過濾後為空或沒有檔案名稱，至少保留檔案名稱（保持原始大小寫）
        if not filtered_path_parts or (filename and filename not in filtered_path_parts):
            if filename:
                # 確保檔案名稱在最後（保持原始大小寫）
                # 使用大小寫不敏感的比較來檢查是否已存在
                filename_lower = filename.lower()
                filtered_path_parts = [p for p in filtered_path_parts if p.lower() != filename_lower]
                filtered_path_parts.append(filename)  # 使用原始文件名（保持大小寫）
        
        os_tags.extend(filtered_path_parts[:max_content_tags])
        return True

    def _extract_all_tags_for_segment(
            self,
            segment: Dict[str, Any],
            segment_index: int,
            segments: List[Dict[str, Any]],
            file_path: Optional[str] = None,
            include_file_path_in_tags: bool = False,
            extract_content_keywords: bool = True,
            max_content_tags: int = 2,
            path_sep: str = None
        ) -> Dict[str, List[str]]:
        """
        為單個 segment 提取所有類型的標籤
        
        Args:
            segment: 當前 segment 字典 
                segment_metadata = segment.get('metadata', {})
                content = segment.get('content', '')
            segment_index: segment 的索引（0-based）
            segments: 所有 segments 列表（用於回溯祖先標題）
            file_path: 文件路徑（可選）
            include_file_path_in_tags: 是否在標籤中包含文件路徑
            extract_content_keywords: 是否提取內容關鍵詞
            max_content_tags: 最大內容標籤數量
            path_sep: 路徑分隔符（預設使用 os.sep）
            
        Returns:
            包含以下鍵的字典：
            - os_tags: 檔案路徑相關標籤
            - struc_tags: 文件結構階層標籤
            - contextual_tags: 脈絡標籤（os_tags + struc_tags，向後兼容）
            - content_tags: 內容標籤（關鍵詞）
        """
        if path_sep is None:
            path_sep = os.sep
        
        segment_metadata = segment.get('metadata', {})
        
        # 1. 提取脈絡標籤，分離為 os_tags 和 struc_tags
        os_tags: List[str] = []  # 檔案路徑相關標籤
        struc_tags: List[str] = []  # 文件結構階層標籤
        
        # 1.1 提取 OS 路徑標籤
        if include_file_path_in_tags and file_path:
            self._extract_os_tags_from_path(
                os_tags=os_tags,
                file_path=file_path,
                path_sep=path_sep,
                max_content_tags=max_content_tags,
            )
        
        # 1.2 提取結構型標籤：回溯祖先標題
        structural_tags = self._extract_contextual_tags_by_backscan(
            segments,
            segment_index,
            level_key='level',
            heading_key='heading'
        )
        struc_tags.extend(structural_tags)
        
        # 1.3 添加當前標題/標識到結構型標籤（如果有）
        current_heading = segment_metadata.get('heading')
        if current_heading and isinstance(current_heading, str) and current_heading.strip():
            struc_tags.append(current_heading.strip())
        
        # 調試信息
        
        # 1.4 過濾掉無意義的標籤
        os_tags = self._filter_meaningless_tags(os_tags)
        struc_tags = self._filter_meaningless_tags(struc_tags)
        
        # 1.5 合併為 contextual_tags（用於向後兼容）
        contextual_tags = os_tags + struc_tags
        
        # 2. 提取內容標籤（關鍵詞）
        content_tags: List[str] = []
        if extract_content_keywords:
            content = segment.get('content', '')
            if content and isinstance(content, str) and content.strip():
                # 只對長度足夠的內容提取關鍵詞（避免不必要的 LLM 調用）
                content_clean = content.strip()
                min_text_length = 10  # 與 _extract_content_keywords 內部的檢查保持一致
                if len(content_clean) >= min_text_length:
                    keywords = self._extract_content_keywords(content)
                    # 限制內容標籤數量
                    content_tags.extend(keywords[:max_content_tags])
        
        return {
            'os_tags': os_tags,
            'struc_tags': struc_tags,
            'contextual_tags': contextual_tags,
            'content_tags': content_tags
        }

    def generate_chunks(
            self, 
            segments: List[Dict[str, Any]], 
            metadata: Dict[str, Any],
            extract_content_keywords: bool = True,
            include_file_path_in_tags: bool = False,
            max_content_tags: int = 2
        ) -> List[Dict[str, Any]]:
        """
        生成符合系統定義的 chunk
        
        Args:
            segments: 分段結果列表；這裡只是llm根據一些文本做的分段，也可以是`[text]`
            metadata: 文檔元數據（包含 document_index_id, document_source 等）
            extract_content_keywords: 是否提取內容關鍵詞（預設 True）
            include_file_path_in_tags: 是否在脈絡標籤中包含文件路徑（預設 False）
            max_content_tags: 內容標籤的最大數量（預設 5）
            
        Returns:
            Chunk 列表，符合 dc_documentChunk 結構
        """
        chunks = []
        
        # 獲取文件路徑（如果需要的話）
        file_path = metadata.get('document_source') or metadata.get('file_path') or metadata.get('source') or ''
        path_sep = os.sep  # 使用系統路徑分隔符

        self.logger.info(f"[generate_chunks] metadata keys: {list(metadata.keys())}, file_path='{file_path}', include_file_path_in_tags={include_file_path_in_tags}")
        self.logger.summary(f"[generate_chunks] 處理 {len(segments)} 個 segments")
        
        # 並行提取關鍵詞（如果啟用）
        keywords_map = {}
        if extract_content_keywords and self.parallel_executor and self.llm_service:
            # 1. 收集所有需要提取關鍵詞的内容
            contents_to_extract = []
            for idx, segment in enumerate(segments):
                content = segment.get('content', '')
                if content and isinstance(content, str) and content.strip():
                    contents_to_extract.append((idx, content))
            
            # 2. 並行提取關鍵詞
            if contents_to_extract:
                self.logger.info(f"[generate_chunks] 並行提取 {len(contents_to_extract)} 個關鍵詞")
                prompts_list = []
                for idx, content in contents_to_extract:
                    prompt_dict = self._build_keyword_prompt(content)
                    if prompt_dict:
                        prompts_list.append((idx, prompt_dict))
                
                if prompts_list:
                    # 並行调用
                    results = self.parallel_executor.batch_chat([p[1] for p in prompts_list])
                    # 解析结果並建立映射
                    for (idx, _), result in zip(prompts_list, results):
                        if result:
                            keywords = self._parse_keywords(result)
                            # 過濾口語用字和無技術成分的用字
                            keywords = self._filter_casual_terms(keywords, max_result_count=max_content_tags)
                            keywords_map[idx] = keywords
                        else:
                            keywords_map[idx] = []
        
        # LLM 去重 multi_prompts（批量）
        try:
            segments = self._dedup_segment_multi_prompts(segments)
        except Exception as e:
            if self.logger:
                self.logger.warning(f"[generate_chunks] multi_prompts 去重失敗: {e}")

        for idx, segment in enumerate(segments, start=1):
            # 生成 ChunkID
            chunk_id = str(uuid.uuid4())
            
            # 獲取 segment 的 metadata
            segment_metadata = segment.get('metadata', {})
            
            # 提取所有類型的標籤（如果已並行提取關鍵詞，则跳過關鍵詞提取）
            segment_idx = idx - 1  # 0-based index
            if extract_content_keywords and segment_idx in keywords_map:
                # 使用预先提取的關鍵詞
                pre_extracted_keywords = keywords_map[segment_idx]
                # 提取其他標签（不包含關鍵詞）
                tags_result = self._extract_all_tags_for_segment(
                    segment=segment,
                    segment_index=segment_idx,
                    segments=segments,
                    file_path=file_path,
                    include_file_path_in_tags=include_file_path_in_tags,
                    extract_content_keywords=False,  # 跳過關鍵詞提取
                    max_content_tags=max_content_tags,
                    path_sep=path_sep
                )
                # 使用预先提取的關鍵詞
                tags_result['content_tags'] = pre_extracted_keywords
            else:
                # 串行提取（降级或未啟用並行）
                tags_result = self._extract_all_tags_for_segment(
                    segment=segment,
                    segment_index=segment_idx,
                    segments=segments,
                    file_path=file_path,
                    include_file_path_in_tags=include_file_path_in_tags,
                    extract_content_keywords=extract_content_keywords,
                    max_content_tags=max_content_tags,
                    path_sep=path_sep
                )
            
            # 構建 chunk 資料
            chunk_metadata = segment_metadata.copy()
            # 將標籤添加到 metadata 中
            chunk_metadata['os_tags'] = tags_result['os_tags']
            chunk_metadata['struc_tags'] = tags_result['struc_tags']
            chunk_metadata['contextual_tags'] = tags_result['contextual_tags']  # 合併版本，向後兼容
            chunk_metadata['content_tags'] = tags_result['content_tags']
            # 確保 document_source 資訊被傳遞
            chunk_metadata['document_source'] = file_path
            
            # 獲取 segment 的 content，嘗試多個可能的鍵
            segment_content = segment.get('content', '')
            if not segment_content:
                # 如果 content 為空，嘗試從 metadata 中獲取
                segment_metadata = segment.get('metadata', {})
                segment_content = segment_metadata.get('content', '')
            
            # 調試：記錄 segment 的內容
            if not segment_content:
                self.logger.warning(f"[generate_chunks] Segment {idx} 的 content 為空，segment keys={list(segment.keys())}, metadata keys={list(segment.get('metadata', {}).keys())}")
            
            # 保留 segment 中的關鍵字段（multi_prompts, call_prompt, tags, title 等）
            # 這些字段需要傳遞給 _create_vector_point_from_chunk 使用
            segment_call_prompt = segment.get('call_prompt', '')
            segment_multi_prompts = segment.get('multi_prompts', None)
            segment_tags = segment.get('tags', [])
            segment_title = segment.get('title', '')
            
            # 如果 multi_prompts 是列表格式，記錄日誌以便調試
            if segment_multi_prompts:
                if isinstance(segment_multi_prompts, list):
                    self.logger.summary(f"[generate_chunks] Segment {idx} 包含 {len(segment_multi_prompts)} 個 multi_prompts（列表格式）")
                elif isinstance(segment_multi_prompts, dict):
                    self.logger.summary(f"[generate_chunks] Segment {idx} 包含 {len(segment_multi_prompts)} 個 multi_prompts（字典格式）")
            
            chunk = {
                'ChunkID': chunk_id,
                'content': segment_content or '',  # 確保至少是空字符串而不是 None
                'DocumentIndexID': metadata.get('document_index_id'),
                # ChunkIndex 不在組合前設置，因為這是組合前的順序，不是文件中的最終順序
                # 組合後會按照在文件中的最終順序重新設置 ChunkIndex
                'ChunkIndex': None,  # 組合前不設置，組合後會重新設置
                'Page': metadata.get('page', None),
                'Language': metadata.get('language', 'default'),
                'RegisteredGroup': metadata.get('registered_group', ''),
                'CreateUser': metadata.get('create_user', 'system'),
                'IsActivate': True,
                'IsDelete': False,
                'source': file_path,
                # 保留 segment 中的關鍵字段，供 _create_vector_point_from_chunk 使用
                'call_prompt': segment_call_prompt,
                'multi_prompts': segment_multi_prompts,
                'tags': segment_tags,
                'title': segment_title,
                # 額外的元數據（包含標籤）
                'metadata': chunk_metadata
            }
            
            chunks.append(chunk)
        
        return chunks

    def generate_chunks_from_docx_structure(
            self,
            docx_structure: List[Dict[str, Any]],
            metadata: Dict[str, Any],
            original_text: str = "",
            include_file_path_in_tags: bool = False,
            use_multi_prompts: bool = True,
            max_contextual_tags: int = 2,
            max_content_tags: int = 5,
            segment_mode: str = "index",
            slide_offset: int = 300,
            enable_chunk_merge: bool = True,
            chunks_per_group: int = 3
        ) -> List[Dict[str, Any]]:
        """
        根據 docx_structure（扁平化階層資料）生成 chunks。

        Args:
            docx_structure: 來自 DocxParser 的扁平化結構資料
            metadata: 文檔元資料（至少需包含 document_source 或 file_path）
        """
        if not docx_structure or not isinstance(docx_structure, list):
            self.logger.warning("[generate_chunks_from_docx_structure] docx_structure 為空或格式錯誤")
            return []

        segments = self._convert_docx_structure_to_segments(docx_structure, metadata)
        if not segments:
            self.logger.warning("[generate_chunks_from_docx_structure] 無法從 docx_structure 生成 segments")
            return []

        chunks = self.generate_chunks(
            segments=segments,
            metadata=metadata,
            extract_content_keywords=True,
            include_file_path_in_tags=include_file_path_in_tags,
            max_content_tags=max_content_tags
        )

        # Docx 結構流程僅需返回 chunks，合併與 multi_prompts 在上層統一處理
        return chunks

    def _convert_docx_structure_to_segments(
            self,
            docx_structure: List[Dict[str, Any]],
            metadata: Dict[str, Any]
        ) -> List[Dict[str, Any]]:
        """
        將 docx_structure 轉換為 ContextParser 可處理的 segments 格式。
        """
        segments: List[Dict[str, Any]] = []
        document_source = metadata.get('document_source') or metadata.get('file_path') or metadata.get('source') or ''
        self.logger.info(f"[_convert_docx_structure_to_segments] metadata keys: {list(metadata.keys())}, document_source={document_source}")

        for idx, item in enumerate(docx_structure):
            if not isinstance(item, dict):
                continue

            content = (item.get('content') or item.get('text') or '').strip()
            if not content:
                continue

            heading = item.get('heading') or item.get('title') or ''
            level = item.get('level', 0)
            indent_level = item.get('indent_level', 0)
            tags = item.get('tags') or item.get('category_tags') or []
            if isinstance(tags, str):
                tags = [tags]

            segment_metadata = {
                'heading': heading,
                'level': level,
                'indent_level': indent_level,
                'tags': tags,
                'category_tags': tags,
                'document_source': document_source,
                'file_path': document_source,
                'index': idx
            }

            segments.append({
                'content': content,
                'metadata': segment_metadata
            })

        return segments

    def build_segments_from_chunks(
            self,
            chunk_list: List[Dict[str, Any]],
            original_text: str,
            use_multi_prompts: bool = True,
            max_contextual_tags: int = 2,
            max_content_tags: int = 2,
            segment_mode: str = "index",
            slide_offset: int = 300
        ) -> List[Dict[str, Any]]:
        """
        將 generate_chunks 的結果轉為 DocumentManager 需要的 segment 結構。
        """
        if not chunk_list:
            return []
        
        processed_segments = []
        text = original_text or ""
        current_index = 0
        
        for chunk in chunk_list:
            content = (chunk.get('Content') or chunk.get('content') or '').strip()
            if not content:
                continue
            
            metadata = chunk.get('metadata', {}) or {}
            heading = metadata.get('heading', '')
            os_tags = metadata.get('os_tags', [])
            struc_tags = metadata.get('struc_tags', [])
            contextual_tags = metadata.get('contextual_tags', [])
            content_tags = metadata.get('content_tags', [])
            
            # 判斷是否是標題 chunk（避免重複標題）
            is_heading_chunk = False
            if heading and heading.strip():
                # 檢查 content 是否等於 heading（標題通常 content 就是 heading）
                content_stripped = content.strip()
                heading_stripped = heading.strip()
                if content_stripped == heading_stripped or (len(content_stripped) < 100 and heading_stripped in content_stripped):
                    is_heading_chunk = True
            
            self.logger.debug(
                f"[build_segments] chunk_index={chunk.get('ChunkIndex') or chunk.get('chunk_order')}, "
                f"heading='{heading}', level={metadata.get('level')}, is_heading_chunk={is_heading_chunk}, "
                f"os_tags={os_tags}, struc_tags={struc_tags}, "
                f"contextual_tags={contextual_tags}, content_tags={content_tags}, "
                f"content_preview='{content[:60]}'"
            )
            
            # 構建 tags 列表（包含所有標籤類型）
            # 如果是標題 chunk，heading 已經在 content 中，不需要重複添加
            tags = []
            if heading and not is_heading_chunk:
                tags.append(heading)
            if os_tags:
                tags.extend(os_tags)
            if struc_tags:
                tags.extend(struc_tags)
            if contextual_tags:
                tags.extend(contextual_tags)
            if content_tags:
                tags.extend(content_tags)
            
            # 去重並過濾系統檔案結構名稱
            # 使用 dict.fromkeys 保持順序的去重
            unique_tags = list(dict.fromkeys(tags))
            # 過濾掉系統檔案結構名稱
            filtered_tags = self._filter_meaningless_tags(unique_tags) if hasattr(self, '_filter_meaningless_tags') else unique_tags
            tags = filtered_tags
            
            if segment_mode == "index" and text:
                start_index = text.find(content, current_index)
                if start_index == -1:
                    start_index = text.find(content)
                if start_index == -1:
                    start_index = current_index
                length = len(content)
                current_index = start_index + length
            else:
                start_index = -1
                length = len(content)
            
            if not isinstance(max_contextual_tags, int):
                max_contextual_tags = 2
            
            # 優先從 chunk 中提取原始的 call_prompt 和 multi_prompts
            original_call_prompt = chunk.get('call_prompt') or metadata.get('call_prompt')
            original_multi_prompts = chunk.get('multi_prompts') or metadata.get('multi_prompts')
            
            multi_prompts: Dict[str, str] = {}
            # 如果找到原始的 multi_prompts，優先使用它；否則重新生成
            if original_multi_prompts:
                # 保留原始的 multi_prompts
                if isinstance(original_multi_prompts, dict):
                    multi_prompts = original_multi_prompts.copy()
                elif isinstance(original_multi_prompts, list):
                    # 列表格式轉換為字典格式（向後兼容）
                    for idx, prompt_value in enumerate(original_multi_prompts):
                        if prompt_value and str(prompt_value).strip():
                            multi_prompts[f'multi_prompt_{idx}'] = str(prompt_value).strip()
                else:
                    # 其他格式，嘗試轉換
                    self.logger.warning(f"[build_segments_from_chunks] 原始 multi_prompts 格式不正確: {type(original_multi_prompts)}")
            elif use_multi_prompts:
                # 重新生成 multi_prompts（基於標籤）
                # 步驟 1: 處理 contextual_tags - 濾除系統結構名稱，計算 jump，產生 max_contextual_tags 個累加標籤
                filtered_contextual = self._filter_meaningless_tags(contextual_tags) if hasattr(self, '_filter_meaningless_tags') else contextual_tags
                filtered_contextual = [tag.strip() for tag in filtered_contextual if isinstance(tag, str) and tag.strip()]
                contextual_prompt_values: List[str] = []
                if filtered_contextual and max_contextual_tags > 0:
                    # 計算 jump（步長），確保最多生成 max_contextual_tags 個累加標籤
                    n = len(filtered_contextual)
                    jump = max(1, math.ceil(n / max_contextual_tags)) if n > max_contextual_tags else 1
                    # 生成累加標籤：從不同位置開始的標籤組合
                    # 例如：['a | b | c | d | e', 'd | e'] (如果 jump=3, max_contextual_tags=2)
                    for i in range(0, n, jump):
                        if i < n:
                            tags_slice = filtered_contextual[i:]
                            if tags_slice:
                                prompt_value = ' | '.join(tags_slice)
                                contextual_prompt_values.append(prompt_value)
                                # 如果已經達到 max_contextual_tags 個，停止
                                if len(contextual_prompt_values) >= max_contextual_tags:
                                    break
                
                # 步驟 2: 處理 content_tags - 每個 content_tag 獨立一個
                content_prompt_values: List[str] = []
                if content_tags:
                    for tag in content_tags:
                        if not tag:
                            continue
                        tag_text = tag.strip() if isinstance(tag, str) else str(tag).strip()
                        if tag_text:
                            content_prompt_values.append(tag_text)
                
                # 步驟 2.5: 處理表格提示詞（如果 segment 包含表格）
                table_prompt_values: List[str] = []
                chunk_metadata = chunk.get('metadata', {})
                if not chunk_metadata:
                    chunk_metadata = metadata
                table_prompts = chunk_metadata.get('table_prompts', [])
                if table_prompts and isinstance(table_prompts, list):
                    for table_prompt in table_prompts:
                        if table_prompt and isinstance(table_prompt, str):
                            table_prompt_text = table_prompt.strip()
                            if table_prompt_text:
                                table_prompt_values.append(table_prompt_text)
                
                # 步驟 3: 組合生成 multi_prompts - 每個 contextual_tags 的累加標籤 × 每個 content_tag = N*M 個
                # 使用 set 來去重 prompt 值
                seen_prompts = set()
                if contextual_prompt_values:
                    for ctx_idx, ctx_prompt in enumerate(contextual_prompt_values):
                        if ctx_prompt not in seen_prompts:
                            seen_prompts.add(ctx_prompt)
                            key = f'contextual_tags_{ctx_idx}'
                            multi_prompts[key] = ctx_prompt
                if content_prompt_values:
                    for content_idx, content_prompt in enumerate(content_prompt_values):
                        if content_prompt not in seen_prompts:
                            seen_prompts.add(content_prompt)
                            key = f'content_tag_{content_idx}'
                            multi_prompts[key] = content_prompt
                # 添加表格提示詞到 multi_prompts
                if table_prompt_values:
                    for table_idx, table_prompt in enumerate(table_prompt_values):
                        if table_prompt not in seen_prompts:
                            seen_prompts.add(table_prompt)
                            key = f'table_prompt_{table_idx}'
                            multi_prompts[key] = table_prompt
            
            # 處理 call_prompt
            call_prompt_value = self._process_call_prompt_for_segment(original_call_prompt, content, max_length=100)
            source_value = chunk.get('source') or metadata.get('document_source', '') or metadata.get('file_path', '')
            
            # 記錄 source 的來源
            self.logger.info(f"[build_segments_from_chunks] Segment {len(processed_segments)+1}: chunk.source={chunk.get('source')}, metadata.document_source={metadata.get('document_source')}, metadata.file_path={metadata.get('file_path')}, 最終 source_value={source_value}")
            if source_value:
                if chunk.get('source'):
                    self.logger.debug(f"[build_segments_from_chunks] Segment {len(processed_segments)+1} 使用 chunk.source: {source_value}")
                elif metadata.get('document_source'):
                    self.logger.debug(f"[build_segments_from_chunks] Segment {len(processed_segments)+1} 使用 metadata.document_source: {source_value}")
                elif metadata.get('file_path'):
                    self.logger.debug(f"[build_segments_from_chunks] Segment {len(processed_segments)+1} 使用 metadata.file_path: {source_value}")
            else:
                self.logger.warning(f"[build_segments_from_chunks] Segment {len(processed_segments)+1} 無法取得 source，chunk.keys()={list(chunk.keys())}, metadata.keys()={list(metadata.keys())}")

            segment = {
                'index': len(processed_segments),
                'content': content,
                'tags': tags,
                'start_index': start_index,
                'length': length,
                'call_prompt': call_prompt_value,
                'os_tags': os_tags,
                'struc_tags': struc_tags,
                'contextual_tags': contextual_tags,
                'content_tags': content_tags,
                'source': source_value,
                'metadata': metadata.copy()
            }
            
            # 將 multi_prompts 添加到 segment（如果是字典，轉換為列表格式）
            if multi_prompts:
                if isinstance(multi_prompts, dict):
                    segment['multi_prompts'] = list(multi_prompts.values())
                elif isinstance(multi_prompts, list):
                    segment['multi_prompts'] = multi_prompts
                else:
                    # 其他格式，嘗試轉換
                    segment['multi_prompts'] = [str(mp) for mp in multi_prompts] if multi_prompts else []
            
            processed_segments.append(segment)
        
        return processed_segments
    
    def merge_chunks_hierarchically(
            self,
            chunks: List[Dict[str, Any]],
            chunks_per_group: int = 3,
            enable_merge: bool = True,
            max_chunk_length: int = 1000
        ) -> List[Dict[str, Any]]:
        """
        Backward compatible wrapper that delegates to _merge_chunks_hierarchically.
        
        Args:
            chunks: 要合併的 chunks 列表
            chunks_per_group: 備用分组大小（當無層級信息時使用）
            enable_merge: 是否啟用合併
            max_chunk_length: 最大 chunk 字數，超過此值不再繼續合併相鄰的最深 level chunks
        """
        return self._merge_chunks_hierarchically(
            chunks=chunks,
            chunks_per_group=chunks_per_group,
            enable_merge=enable_merge,
            max_chunk_length=max_chunk_length
        )
    
    def _merge_chunks_hierarchically(self, chunks, chunks_per_group=3, enable_merge=True, max_chunk_length=1000):
        """
        基於層級结构合併 chunks
        
        優先策略：按最深層級合併
        - 找出最深層級（level 最大）的 chunks
        - 相鄰的相同最深層級 chunks 合併，往上追溯較淺的 level
        - 統計字數，如果 < max_chunk_length，繼續合併下一個相鄰的最深 level chunk
        - 標題追溯到 level=0 的層級
        
        備用策略：當無層級信息時，使用 chunks_per_group
        
        Args:
            chunks: 要合併的 chunks 列表
            chunks_per_group: 備用分组大小（當無層級信息時使用）
            enable_merge: 是否啟用合併
            max_chunk_length: 最大 chunk 字數，超過此值不再繼續合併相鄰的最深 level chunks
            
        Returns:
            合併后的 chunks 列表
        """
        import copy
        
        if not enable_merge or not chunks:
            # 沒有合併，但仍需要设置 ChunkIndex
            final_chunks = []
            for idx, chunk in enumerate(chunks):
                final_chunk = copy.deepcopy(chunk) if isinstance(chunk, dict) else chunk
                if isinstance(final_chunk, dict):
                    final_chunk['ChunkIndex'] = idx + 1
                    final_chunk['chunk_order'] = idx + 1
                final_chunks.append(final_chunk)
            return final_chunks
        
        # 1. 檢查是否有層級信息
        has_level_info = any(
            chunk.get('level') is not None or 
            chunk.get('metadata', {}).get('level') is not None 
            for chunk in chunks
        )
        self.logger.debug(f"[merge] start, chunks={len(chunks)}, has_level_info={has_level_info}, "
                   f"chunks_per_group={chunks_per_group}, max_chunk_length={max_chunk_length}")
        
        if has_level_info:
            # 基於層級合併
            self.logger.info(f"[merge] 檢測到層級信息，使用層級合併策略，共 {len(chunks)} 個 chunks，max_chunk_length={max_chunk_length}")
            return self._merge_by_hierarchy(chunks, max_chunk_length=max_chunk_length)
        else:
            # 備用：使用 chunks_per_group
            self.logger.info(f"[merge] 無層級信息，使用分组合併策略 (chunks_per_group={chunks_per_group})，共 {len(chunks)} 個 chunks")
            return self._merge_by_group_size(chunks, chunks_per_group)

    def _merge_by_hierarchy(self, chunks, max_chunk_length=1000):
        """
        按最深層級結構合併 chunks，符合 main.txt 需求規範
        
        策略（符合 main.txt 需求）：
        1. 深層優先回溯找標題、組成最小群組：從最深層 level = max_level 開始，對每個 chunk 向前回溯，
           尋找最近的上層有 heading 的 chunk，回溯不超過 level=0
        2. 若一路回到 level=0 仍無 heading，則使用 level=0 的第一句作為標題
        3. 由深到淺逐層收攏：完成最深層分組後，將每個群組視為「單位」，再處理上一層（n-1）、
           再上一層（n-2）…直到 level=0
        4. 保持順序與排序鍵：以群組內最小原始索引作為排序鍵
        5. 標題缺失的群組不要登記，並且要回傳缺失資訊
        
        Args:
            chunks: 要合併的 chunks 列表
            max_chunk_length: 最大 chunk 字數，超過此值不再繼續合併相鄰的最深 level chunks
            
        Returns:
            合併后的 chunks 列表
        """
        import copy
        
        def _get_chunk_level(chunk):
            """獲取 chunk 的 level"""
            level = chunk.get('level')
            if level is None:
                level = chunk.get('metadata', {}).get('level')
            return level if level is not None else float('inf')
        
        def _get_chunk_content(chunk):
            """獲取 chunk 的內容"""
            content = chunk.get('content', '') or chunk.get('text', '') or chunk.get('Content', '')
            if not content:
                metadata = chunk.get('metadata', {})
                content = metadata.get('content', '') or metadata.get('text', '')
            return content if content else ''
        
        def _get_chunk_content_length(chunk):
            """獲取 chunk 的内容長度"""
            return len(_get_chunk_content(chunk))
        
        def _get_chunk_heading(chunk):
            """獲取 chunk 的標題"""
            heading = chunk.get('heading', '') or chunk.get('title', '')
            if not heading:
                metadata = chunk.get('metadata', {})
                heading = metadata.get('heading', '') or metadata.get('title', '')
            return heading if heading else ''
        
        def _is_heading_chunk(chunk):
            """判斷 chunk 是否是標題（有 heading 且不是內文）"""
            heading = _get_chunk_heading(chunk)
            if not heading or not heading.strip():
                return False
            
            # 檢查 content 是否等於 heading（標題通常 content 就是 heading）
            content = _get_chunk_content(chunk)
            
            # 如果 content 為空或很短，且等於 heading，則認為是標題
            if content and content.strip():
                content_stripped = content.strip()
                heading_stripped = heading.strip()
                # 如果 content 等於 heading，或者 content 很短（可能是標題後沒有內容）
                if content_stripped == heading_stripped or (len(content_stripped) < 100 and heading_stripped in content_stripped):
                    return True
            else:
                # content 為空但有 heading，認為是標題
                return True
            
            return False
        
        def _find_heading_backward(chunk_idx, target_level):
            """
            從 chunk_idx 往前回溯，尋找 level <= target_level 且有 heading 的 chunk
            回溯不超過 level=0
            
            Returns:
                (heading_idx, heading): 找到的標題 chunk 索引和標題，如果沒找到則返回 (-1, '')
            """
            for j in range(chunk_idx, -1, -1):
                prev_chunk = chunks[j]
                prev_level = _get_chunk_level(prev_chunk)
                
                # 如果遇到更深的 level，停止追溯
                if prev_level > target_level:
                    continue
                
                # 如果找到目標層級或更淺的層級
                if prev_level <= target_level:
                    heading = _get_chunk_heading(prev_chunk)
                    if heading:
                        return (j, heading)
                    # 如果找到 level=0 但沒有 heading，記錄這個位置用於 fallback
                    if prev_level == 0:
                        return (j, '')  # 返回空標題，表示需要 fallback
            
            return (-1, '')
        
        def _get_first_sentence_from_level0(chunk_idx):
            """
            從 level=0 的 chunk 中提取第一句作為標題
            
            Args:
                chunk_idx: level=0 的 chunk 索引
                
            Returns:
                第一句內容，如果無法提取則返回空字串
            """
            if chunk_idx < 0 or chunk_idx >= len(chunks):
                return ''
            
            chunk = chunks[chunk_idx]
            content = _get_chunk_content(chunk)
            if not content:
                return ''
            
            # 提取第一句：第一行，或第一個句號/問號/驚嘆號之前的內容
            first_line = content.split('\n')[0].strip()
            if not first_line:
                return ''
            
            # 找第一個句號、問號、驚嘆號
            for sep in ['。', '.', '？', '?', '！', '!']:
                idx = first_line.find(sep)
                if idx > 0:
                    return first_line[:idx].strip()
            
            # 如果沒找到分隔符，返回前50字符
            return first_line[:50].strip() if len(first_line) > 50 else first_line
        
        # 1. 找出最深層級（level 最大，非 None）
        max_level = -1
        for chunk in chunks:
            level = _get_chunk_level(chunk)
            if level != float('inf') and level > max_level:
                max_level = level
        
        # 邊界情况：如果沒有找到有效的 level，退回到按组大小合併
        if max_level == -1:
            self.logger.warning(f"[merge_by_hierarchy] 未找到有效的層級信息，退回到分组合併")
            return self._merge_by_group_size(chunks, chunks_per_group=3)
        
        # 統計所有 chunks 的 level 分布
        level_stats = {}
        for idx, chunk in enumerate(chunks):
            level = _get_chunk_level(chunk)
            if level not in level_stats:
                level_stats[level] = []
            heading = _get_chunk_heading(chunk)
            level_stats[level].append((idx, heading or f"chunk_{idx}"))
        
        self.logger.info(f"[merge_by_hierarchy] 最深層級 (max_level) = {max_level}, max_chunk_length = {max_chunk_length}")
        self.logger.info(f"[merge_by_hierarchy] Level 分布: {dict([(k, len(v)) for k, v in level_stats.items()])}")
        
        for level in sorted(level_stats.keys()):
            if level != float('inf'):
                chunks_info = level_stats[level][:5]  # 只顯示前5個
                self.logger.debug(f"[merge_by_hierarchy] Level {level}: {chunks_info}")
                if len(level_stats[level]) > 5:
                    self.logger.debug(f"[merge_by_hierarchy] Level {level}: ... 還有 {len(level_stats[level]) - 5} 個")
        
        # 2. 異常情况檢測：如果 level=0 過多且總字數過少，说明層級識别失败
        level0_count = sum(1 for chunk in chunks if _get_chunk_level(chunk) == 0)
        total_length = sum(_get_chunk_content_length(chunk) for chunk in chunks)
        level0_ratio = level0_count / len(chunks) if len(chunks) > 0 else 0
        
        # 判定条件：level=0 占比 > 50% 且總字數 < 500
        is_level_detection_failed = (level0_ratio > 0.5 and total_length < 500)
        
        if is_level_detection_failed:
            self.logger.warning(
                f"[merge_by_hierarchy] 層級識别異常：level=0 占比 {level0_ratio:.1%} ({level0_count}/{len(chunks)})，"
                f"總字數 {total_length}，改用按字數順序合併"
            )
            return self._merge_by_content_length(chunks, max_chunk_length)
        
        # 2.5. 檢查最深層級是否過少（防止整份文件被最深層拉走，只剩少數段）
        max_level_chunks = [chunk for chunk in chunks if _get_chunk_level(chunk) == max_level]
        max_level_count = len(max_level_chunks)
        max_level_ratio = max_level_count / len(chunks) if len(chunks) > 0 else 0
        
        # 判定条件：最深層級 chunk 數量 < 3 個，或占比 < 5%
        is_deep_level_scarce = (max_level_count < 3) or (max_level_ratio < 0.05)
        
        if is_deep_level_scarce:
            self.logger.warning(
                f"[merge_by_hierarchy] 最深層級過少：max_level={max_level} 的 chunks 只有 {max_level_count} 個（占比 {max_level_ratio:.1%}），"
                f"改用按字數順序合併，避免整份文件被最深層拉走只剩少數段"
            )
            return self._merge_by_content_length(chunks, max_chunk_length)
        
        # 3. 由深到淺逐層收攏：從最深層開始，逐層處理直到 level=0
        # 每個群組記錄：包含的原始 chunk 索引列表、標題、最小索引（用於排序）
        groups_by_level = {}  # {level: [group1, group2, ...]}，每個 group 包含原始 chunk 索引列表
        
        # 3.1 處理最深層級（max_level）
        processed_indices = set()
        current_level_groups = []
        
        for i in range(len(chunks)):
            if i in processed_indices:
                continue
            
            chunk = chunks[i]
            chunk_level = _get_chunk_level(chunk)
            
            # 只處理最深 level 的 chunks
            if chunk_level != max_level:
                continue
            
            # 開始一個新的合併組
            current_group_indices = []
            current_length = 0
            
            # 往前回溯，尋找最近的上層有 heading 的 chunk，回溯不超過 level=0
            # 從最深層開始，逐層往上找（max_level-1, max_level-2, ..., 0）
            found_heading_idx = -1
            found_heading = ''
            found_level0_idx = -1
            
            # 特殊處理：當 max_level = 0 時，直接檢查當前 chunk 的標題或使用 fallback
            if max_level == 0:
                current_heading = _get_chunk_heading(chunk)
                if current_heading:
                    found_heading_idx = i
                    found_heading = current_heading
                    self.logger.debug(f"[merge_by_hierarchy] Level 0 chunk {i} 找到標題: '{found_heading}'")
                else:
                    # Level 0 但沒有標題，使用第一句作為標題
                    found_heading = _get_first_sentence_from_level0(i)
                    found_heading_idx = i
                    if found_heading:
                        self.logger.debug(f"[merge_by_hierarchy] Level 0 chunk {i} 沒有標題，使用第一句作為標題: '{found_heading}'")
                    else:
                        # 如果連第一句都提取不到，使用 content 前50字符作為標題
                        content = _get_chunk_content(chunk)
                        if content:
                            found_heading = content[:50].strip()
                            self.logger.debug(f"[merge_by_hierarchy] Level 0 chunk {i} 使用 content 前50字符作為標題: '{found_heading}'")
                        else:
                            self.logger.warning(f"[merge_by_hierarchy] Level 0 chunk {i} 完全沒有標題和內容")
            else:
                # 非 Level 0 的情況，執行回溯邏輯
                for target_level in range(max_level - 1, -1, -1):
                    heading_idx, heading = _find_heading_backward(i, target_level)
                    if heading_idx != -1:
                        if heading:  # 找到有標題的 chunk
                            found_heading_idx = heading_idx
                            found_heading = heading
                            break
                        elif target_level == 0:  # 找到 level=0 但沒有標題，記錄用於 fallback
                            found_level0_idx = heading_idx
                
                # 如果沒找到標題，且找到了 level=0 的 chunk，使用第一句作為標題
                if not found_heading and found_level0_idx != -1:
                    found_heading = _get_first_sentence_from_level0(found_level0_idx)
                    found_heading_idx = found_level0_idx
                    self.logger.debug(f"[merge_by_hierarchy] 使用 level=0 的第一句作為標題: '{found_heading}'")
            
            # 確定群組範圍：從找到的標題 chunk（或當前 chunk）到當前最深 level chunk
            start_idx = found_heading_idx if found_heading_idx != -1 else i
            
            # 收集從 start_idx 到 i 之間的所有 chunks（保持原文順序）
            for j in range(start_idx, i + 1):
                if j not in processed_indices:
                    current_group_indices.append(j)
                    current_length += _get_chunk_content_length(chunks[j])
            
            # 繼續往後合併相鄰的相同最深 level 的 chunks（如果字數未超過限制）
            j = i + 1
            while j < len(chunks):
                next_chunk = chunks[j]
                next_level = _get_chunk_level(next_chunk)
                
                if next_level == max_level and j not in processed_indices:
                    next_length = _get_chunk_content_length(next_chunk)
                    
                    # 檢查是否屬於同一個較淺層級標題下
                    same_shallow_level = False
                    for k in range(j, -1, -1):
                        prev_chunk = chunks[k]
                        prev_level = _get_chunk_level(prev_chunk)
                        if prev_level < max_level:
                            prev_heading = _get_chunk_heading(prev_chunk)
                            if prev_heading and found_heading_idx != -1 and k == found_heading_idx:
                                same_shallow_level = True
                            break
                        elif prev_level > max_level:
                            break
                    
                    if same_shallow_level and current_length + next_length <= max_chunk_length:
                        # 字數未超過限制，且屬於同一個較淺層級標題下，繼續合併
                        current_group_indices.append(j)
                        current_length += next_length
                        processed_indices.add(j)
                        j += 1
                    else:
                        break
                else:
                    break
            
            # 標記所有已處理的索引
            for idx in current_group_indices:
                processed_indices.add(idx)
            
            # 記錄群組（包含原始 chunk 索引列表、標題、最小索引）
            if current_group_indices:
                min_index = min(current_group_indices)
                current_level_groups.append({
                    'indices': current_group_indices,
                    'heading': found_heading,
                    'min_index': min_index,
                    'level': max_level
                })
        
        groups_by_level[max_level] = current_level_groups
        self.logger.info(f"[merge_by_hierarchy] Level {max_level} 分組完成: {len(current_level_groups)} 個群組")
        
        # 3.2 由深到淺逐層收攏：處理上一層（max_level-1, max_level-2, ..., 0）
        for current_level in range(max_level - 1, -1, -1):
            # 找出當前層級中尚未被包含在任何群組中的 chunks
            unprocessed_indices = set()
            for i, chunk in enumerate(chunks):
                chunk_level = _get_chunk_level(chunk)
                if chunk_level == current_level:
                    # 檢查是否已經被包含在更深的層級群組中
                    is_processed = False
                    for level in range(current_level + 1, max_level + 1):
                        if level in groups_by_level:
                            for group in groups_by_level[level]:
                                if i in group['indices']:
                                    is_processed = True
                                    break
                            if is_processed:
                                break
                    if not is_processed:
                        unprocessed_indices.add(i)
            
            if not unprocessed_indices:
                groups_by_level[current_level] = []
                continue
            
            # 對當前層級的未處理 chunks 進行分組
            current_level_groups = []
            processed_current_level = set()
            
            for i in sorted(unprocessed_indices):
                if i in processed_current_level:
                    continue
                
                chunk = chunks[i]
                chunk_level = _get_chunk_level(chunk)
                
                if chunk_level != current_level:
                    continue
                
                # 開始一個新的合併組
                current_group_indices = []
                current_length = 0
                
                # 往前回溯，尋找最近的上層有 heading 的 chunk，回溯不超過 level=0
                found_heading_idx = -1
                found_heading = ''
                found_level0_idx = -1
                
                # 如果當前 chunk 本身就是 level=0 且有標題，直接使用它
                if current_level == 0:
                    current_heading = _get_chunk_heading(chunk)
                    if current_heading:
                        found_heading_idx = i
                        found_heading = current_heading
                    else:
                        # Level 0 但沒有標題，使用第一句作為標題
                        found_heading = _get_first_sentence_from_level0(i)
                        found_heading_idx = i
                        self.logger.debug(f"[merge_by_hierarchy] Level 0 chunk {i} 沒有標題，使用第一句作為標題: '{found_heading}'")
                else:
                    # 對於非 Level 0 的層級，往前回溯尋找標題
                    for target_level in range(current_level - 1, -1, -1):
                        heading_idx, heading = _find_heading_backward(i, target_level)
                        if heading_idx != -1:
                            if heading:  # 找到有標題的 chunk
                                found_heading_idx = heading_idx
                                found_heading = heading
                                break
                            elif target_level == 0:  # 找到 level=0 但沒有標題，記錄用於 fallback
                                found_level0_idx = heading_idx
                    
                    # 如果沒找到標題，且找到了 level=0 的 chunk，使用第一句作為標題
                    if not found_heading and found_level0_idx != -1:
                        found_heading = _get_first_sentence_from_level0(found_level0_idx)
                        found_heading_idx = found_level0_idx
                        self.logger.debug(f"[merge_by_hierarchy] Level {current_level} 使用 level=0 的第一句作為標題: '{found_heading}'")
                
                # 確定群組範圍：從找到的標題 chunk（或當前 chunk）到當前 level chunk
                start_idx = found_heading_idx if found_heading_idx != -1 else i
                
                # 收集從 start_idx 到 i 之間的所有 chunks（保持原文順序）
                for j in range(start_idx, i + 1):
                    if j not in processed_current_level:
                        prev_chunk_level = _get_chunk_level(chunks[j])
                        # 只包含當前層級或更淺層級的 chunks
                        if prev_chunk_level <= current_level:
                            current_group_indices.append(j)
                            current_length += _get_chunk_content_length(chunks[j])
                
                # 繼續往後合併相鄰的相同層級的 chunks（如果字數未超過限制）
                j = i + 1
                while j < len(chunks):
                    next_chunk = chunks[j]
                    next_level = _get_chunk_level(next_chunk)
                    
                    if next_level == current_level and j not in processed_current_level:
                        next_length = _get_chunk_content_length(next_chunk)
                        
                        # 檢查是否屬於同一個較淺層級標題下
                        same_shallow_level = False
                        for k in range(j, -1, -1):
                            prev_chunk = chunks[k]
                            prev_level = _get_chunk_level(prev_chunk)
                            if prev_level < current_level:
                                prev_heading = _get_chunk_heading(prev_chunk)
                                if prev_heading and found_heading_idx != -1 and k == found_heading_idx:
                                    same_shallow_level = True
                                break
                            elif prev_level > current_level:
                                break
                        
                        if same_shallow_level and current_length + next_length <= max_chunk_length:
                            # 字數未超過限制，且屬於同一個較淺層級標題下，繼續合併
                            current_group_indices.append(j)
                            current_length += next_length
                            processed_current_level.add(j)
                            j += 1
                        else:
                            break
                    elif next_level < current_level:
                        # 遇到更淺的層級，停止
                        break
                    else:
                        j += 1
                
                # 標記所有已處理的索引
                for idx in current_group_indices:
                    processed_current_level.add(idx)
                
                # 記錄群組
                if current_group_indices:
                    min_index = min(current_group_indices)
                    current_level_groups.append({
                        'indices': current_group_indices,
                        'heading': found_heading,
                        'min_index': min_index,
                        'level': current_level
                    })
            
            groups_by_level[current_level] = current_level_groups
            self.logger.info(f"[merge_by_hierarchy] Level {current_level} 分組完成: {len(current_level_groups)} 個群組")
        
        # 4. 合併所有層級的群組，檢查標題缺失，並按最小原始索引排序
        all_groups = []
        missing_heading_groups = []  # 記錄標題缺失的群組資訊
        
        # 收集所有層級的群組
        for level in range(max_level, -1, -1):
            if level in groups_by_level:
                for group in groups_by_level[level]:
                    all_groups.append(group)
        
        # 檢查標題缺失的群組
        valid_groups = []
        for group in all_groups:
            if not group['heading'] or not group['heading'].strip():
                # 標題缺失，記錄資訊但不登記
                missing_heading_groups.append({
                    'level': group['level'],
                    'indices': group['indices'],
                    'min_index': group['min_index']
                })
                self.logger.warning(
                    f"[merge_by_hierarchy] 標題缺失的群組（level={group['level']}, "
                    f"chunks={group['indices']}, min_index={group['min_index']}）已跳過"
                )
            else:
                valid_groups.append(group)
        
        # 以群組內最小原始索引作為排序鍵，確保群組在最終結果中仍符合原文流
        valid_groups.sort(key=lambda g: g['min_index'])
        
        self.logger.info(f"[merge_by_hierarchy] 有效群組數: {len(valid_groups)}, 標題缺失群組數: {len(missing_heading_groups)}")
        
        # 4.5. 長度門檻判斷：使用 buffer 機制累積群組長度，達到 max_chunk_length 時才封板
        final_groups = []  # 最終要合併的群組列表
        buffer_groups = []  # 暫存的群組 buffer
        current_length = 0  # 當前 buffer 累積的長度
        
        def _calculate_group_length(group_indices):
            """計算群組的內容長度"""
            total_length = 0
            for idx in group_indices:
                if idx >= 0 and idx < len(chunks):
                    total_length += _get_chunk_content_length(chunks[idx])
            return total_length
        
        def _flush_buffer():
            """將 buffer 中的群組合併成一個最終群組"""
            nonlocal current_length
            if not buffer_groups:
                return
            
            # 收集 buffer 中所有群組的所有 chunk 索引
            all_indices = []
            for group in buffer_groups:
                all_indices.extend(group['indices'])
            
            # 去重並排序（保持原文順序）
            all_indices = sorted(set(all_indices))
            
            # 提取標題：按照書寫順序找 level=0 標題
            heading = self._extract_group_heading_from_buffer(buffer_groups, chunks, _get_chunk_level, _get_chunk_heading)
            
            # 計算最小索引（用於排序）
            min_index = min(all_indices) if all_indices else 0
            
            # 創建最終群組
            final_group = {
                'indices': all_indices,
                'heading': heading,
                'min_index': min_index,
                'level': buffer_groups[0]['level'] if buffer_groups else 0
            }
            
            final_groups.append(final_group)
            
            self.logger.debug(
                f"[merge_by_hierarchy] Buffer 封板: {len(buffer_groups)} 個群組合併成 1 個最終群組, "
                f"總長度={current_length}, 標題='{heading}', 索引範圍={all_indices}"
            )
            
            # 清空 buffer
            buffer_groups.clear()
            current_length = 0
        
        # 遍歷 valid_groups，累積長度並判斷是否封板
        for group in valid_groups:
            # 計算當前群組的長度
            group_length = _calculate_group_length(group['indices'])
            
            # 將當前群組加入 buffer
            buffer_groups.append(group)
            current_length += group_length
            
            # 檢查是否達到門檻
            if current_length >= max_chunk_length:
                # 達到門檻，封板
                _flush_buffer()
            # 如果未達門檻，繼續累積到下一個群組
        
        # 遍歷完所有群組後，flush 剩餘的 buffer
        if buffer_groups:
            _flush_buffer()
        
        self.logger.info(f"[merge_by_hierarchy] 長度門檻判斷完成: {len(valid_groups)} 個群組 -> {len(final_groups)} 個最終群組")
        
        # 5. 生成最終的合併結果
        merged_chunks = []
        for group_idx, group in enumerate(final_groups):
            # 從原始 chunks 中提取群組的 chunks
            group_chunks = [chunks[i] for i in group['indices']]
            
            # 合併群組內容
            merged_chunk = self._merge_chunk_group_content_only(group_chunks, len(merged_chunks))
            
            # 設置標題
            if 'metadata' not in merged_chunk:
                merged_chunk['metadata'] = {}
            merged_chunk['metadata']['heading'] = group['heading']
            merged_chunk['metadata']['shallow_heading'] = group['heading']
            
            # 記錄群組的原始索引範圍（用於除錯）
            merged_chunk['metadata']['original_indices'] = group['indices']
            merged_chunk['metadata']['min_original_index'] = group['min_index']
            
            merged_chunks.append(merged_chunk)
            
            # 詳細顯示合併的 chunks 信息
            group_info = []
            for idx in group['indices']:
                c = chunks[idx]
                c_level = _get_chunk_level(c)
                c_heading = _get_chunk_heading(c)
                group_info.append(f"index{idx}:level{c_level}:{c_heading or 'content'}")
            
            self.logger.info(
                f"[merge_by_hierarchy] 合併组 {len(merged_chunks)}: {len(group_chunks)} 個 chunks, "
                f"標題='{group['heading']}', 原始索引範圍={group['indices']}, 最小索引={group['min_index']}"
            )
            self.logger.debug(f"[merge_by_hierarchy] 合併组 {len(merged_chunks)} 詳情: {group_info}")
        
        self.logger.info(f"[merge_by_hierarchy] 合併完成: {len(chunks)} -> {len(merged_chunks)} 個 chunks")
        
        # 如果有標題缺失的群組，記錄到 metadata 中供檢查
        if missing_heading_groups:
            self.logger.warning(
                f"[merge_by_hierarchy] 共有 {len(missing_heading_groups)} 個標題缺失的群組未被登記"
            )
            # 可以將 missing_heading_groups 記錄到某個地方供後續檢查
            # 這裡暫時只記錄到日誌中
        
        # 3. 並行生成 call_prompts
        if self.parallel_executor and merged_chunks:
            self.logger.info(f"[merge_by_hierarchy] 並行生成 {len(merged_chunks)} 個 call_prompts")
            prompts_list = []
            valid_indices = []
            for idx, chunk in enumerate(merged_chunks):
                prompt_dict = self._build_base_prompt_input(chunk)
                if prompt_dict:
                    prompts_list.append(prompt_dict)
                    valid_indices.append(idx)
                else:
                    # 内容過短，使用 fallback
                    merged_metadata = chunk.get('metadata', {})
                    heading = merged_metadata.get('heading', '')
                    struc_tags = merged_metadata.get('struc_tags', [])
                    contextual_tags = merged_metadata.get('contextual_tags', [])
                    content = chunk.get('content', '')
                    fallback_prompt = _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                    merged_metadata['call_prompt'] = fallback_prompt
            
            if prompts_list:
                # 並行调用
                results = self.parallel_executor.batch_chat(prompts_list)
                # 分配结果
                for (idx, prompt_dict), result in zip([(valid_indices[i], prompts_list[i]) for i in range(len(valid_indices))], results):
                    merged_metadata = merged_chunks[idx].get('metadata', {})
                    if result:
                        # 清理响應
                        call_prompt = result.strip().strip('"\'').replace('\n', ' ').replace('\r', ' ')
                        # 移除 LLM 結束標記 <|eot_id|>
                        if '<|eot_id|>' in call_prompt:
                            call_prompt = call_prompt.split('<|eot_id|>')[0].strip()
                        if len(call_prompt) > 200:
                            call_prompt = call_prompt[:200].strip()
                        # 檢測錯誤消息
                        error_patterns = self.config.get('llm_error', {}).get('patterns', [])
                        if any(pattern in call_prompt for pattern in error_patterns):
                            # 使用 fallback
                            heading = merged_metadata.get('heading', '')
                            struc_tags = merged_metadata.get('struc_tags', [])
                            contextual_tags = merged_metadata.get('contextual_tags', [])
                            content = merged_chunks[idx].get('content', '')
                            call_prompt = _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                        
                        # 檢查是否需要過濾（長度 <= threshold）
                        if self.filter_casual_terms_enabled and len(call_prompt) <= self.filter_casual_terms_threshold:
                            if self._should_filter_single_term(call_prompt):
                                # 應該過濾，使用 fallback
                                heading = merged_metadata.get('heading', '')
                                struc_tags = merged_metadata.get('struc_tags', [])
                                contextual_tags = merged_metadata.get('contextual_tags', [])
                                content = merged_chunks[idx].get('content', '')
                                merged_metadata['call_prompt'] = _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                            else:
                                # 不需要過濾，保留原 call_prompt
                                merged_metadata['call_prompt'] = call_prompt
                        else:
                            merged_metadata['call_prompt'] = call_prompt
                    else:
                        # 使用 fallback
                        heading = merged_metadata.get('heading', '')
                        struc_tags = merged_metadata.get('struc_tags', [])
                        contextual_tags = merged_metadata.get('contextual_tags', [])
                        content = merged_chunks[idx].get('content', '')
                        merged_metadata['call_prompt'] = _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                
                # 並行判斷需要過濾的 call_prompts（已在上面處理，這裡不需要額外處理）
        else:
            # 降级：串行生成 call_prompts
            for chunk in merged_chunks:
                merged_metadata = chunk.get('metadata', {})
                if not merged_metadata.get('call_prompt'):
                    # 如果还沒有 call_prompt，生成它
                    heading = merged_metadata.get('heading', '')
                    os_tags = merged_metadata.get('os_tags', [])
                    struc_tags = merged_metadata.get('struc_tags', [])
                    contextual_tags = merged_metadata.get('contextual_tags', [])
                    content_tags = merged_metadata.get('content_tags', [])
                    content = chunk.get('content', '')
                    all_tags = os_tags + struc_tags + content_tags
                    unique_all_tags = list(dict.fromkeys(all_tags))
                    
                    call_prompt = _generate_base_prompt_with_llm(
                        content=content,
                        heading=heading,
                        tags=unique_all_tags,
                        os_tags=os_tags,
                        struc_tags=struc_tags,
                        contextual_tags=contextual_tags,
                        content_tags=content_tags,
                        llm_config=self.config.get('llm', {})
                    )
                    
                    # 檢查是否需要過濾（長度 <= threshold）
                    if self.filter_casual_terms_enabled and len(call_prompt) <= self.filter_casual_terms_threshold:
                        if self._should_filter_single_term(call_prompt):
                            # 應該過濾，使用 fallback
                            merged_metadata['call_prompt'] = _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                        else:
                            # 不需要過濾，保留原 call_prompt
                            merged_metadata['call_prompt'] = call_prompt
                    else:
                        merged_metadata['call_prompt'] = call_prompt
        
        # 4. 设置 ChunkIndex（按合併后的順序）
        for idx, chunk in enumerate(merged_chunks):
            if isinstance(chunk, dict):
                chunk['ChunkIndex'] = idx + 1
                chunk['chunk_order'] = idx + 1
        
        return merged_chunks

    def _merge_by_group_size(self, chunks, chunks_per_group):
        """
        備用方案：按 chunks_per_group 分组合併
        
        當 chunks 沒有層級信息時使用此方法
        
        Args:
            chunks: 要合併的 chunks 列表
            chunks_per_group: 每组的 chunk 数量
            
        Returns:
            合併后的 chunks 列表
        """
        import copy
        
        if chunks_per_group <= 1:
            # 不合併，但仍需要设置 ChunkIndex
            final_chunks = []
            for idx, chunk in enumerate(chunks):
                final_chunk = copy.deepcopy(chunk) if isinstance(chunk, dict) else chunk
                if isinstance(final_chunk, dict):
                    final_chunk['ChunkIndex'] = idx + 1
                    final_chunk['chunk_order'] = idx + 1
                final_chunks.append(final_chunk)
            return final_chunks
        
        merged_chunks = []
        for group_start in range(0, len(chunks), chunks_per_group):
            group_end = min(group_start + chunks_per_group, len(chunks))
            group_chunks = chunks[group_start:group_end]
            
            if len(group_chunks) == 1:
                # 单個 chunk，直接添加
                single_chunk = copy.deepcopy(group_chunks[0])
                single_chunk['ChunkIndex'] = len(merged_chunks) + 1
                single_chunk['chunk_order'] = len(merged_chunks) + 1
                merged_chunks.append(single_chunk)
            else:
                # 多個 chunks，合併
                merged_chunk = self._merge_chunk_group_content_only(group_chunks, len(merged_chunks))
                merged_chunks.append(merged_chunk)
            
            self.logger.debug(f"[merge_by_group_size] 组 {len(merged_chunks)}: {len(group_chunks)} 個 chunks")
        
        self.logger.info(f"[merge_by_group_size] 合併完成: {len(chunks)} -> {len(merged_chunks)} 個 chunks")
        return merged_chunks
    
    def _merge_by_content_length(self, chunks, max_chunk_length=500):
        """
        按内容字數順序合併（用於層級識别失败的情况）
        
        策略：
        1. 從第一個 chunk 开始，累加字數
        2. 直到达到 max_chunk_length，开始新组
        3. 繼續合併直到所有 chunks 處理完
        
        Args:
            chunks: 要合併的 chunks 列表
            max_chunk_length: 最大 chunk 字數
            
        Returns:
            合併后的 chunks 列表
        """
        import copy
        
        def _get_chunk_content_length(chunk):
            """獲取 chunk 的内容長度"""
            content = chunk.get('content', '') or chunk.get('text', '') or chunk.get('Content', '')
            if not content:
                metadata = chunk.get('metadata', {})
                content = metadata.get('content', '') or metadata.get('text', '')
            return len(content) if content else 0
        
        if not chunks:
            return []
        
        self.logger.info(f"[merge_by_content_length] 开始按字數合併: {len(chunks)} 個 chunks, max_chunk_length={max_chunk_length}")
        
        merged_chunks = []
        current_group = []
        current_length = 0
        
        for chunk in chunks:
            chunk_length = _get_chunk_content_length(chunk)
            
            # 如果加入當前 chunk 会超過限制，且當前组不為空，则先合併當前组
            if current_length + chunk_length > max_chunk_length and current_group:
                merged_chunk = self._merge_chunk_group_content_only(current_group, len(merged_chunks))
                merged_chunks.append(merged_chunk)
                self.logger.debug(f"[merge_by_content_length] 合併组 {len(merged_chunks)}: {len(current_group)} 個 chunks, 字數={current_length}")
                
                # 开始新组
                current_group = [chunk]
                current_length = chunk_length
            else:
                # 繼續累加到當前组
                current_group.append(chunk)
                current_length += chunk_length
        
        # 處理最后一组
        if current_group:
            merged_chunk = self._merge_chunk_group_content_only(current_group, len(merged_chunks))
            merged_chunks.append(merged_chunk)
            self.logger.debug(f"[merge_by_content_length] 合併组 {len(merged_chunks)}: {len(current_group)} 個 chunks, 字數={current_length}")
        
        self.logger.info(f"[merge_by_content_length] 合併完成: {len(chunks)} -> {len(merged_chunks)} 個 chunks")
        
        # 並行生成 call_prompts（与 _merge_by_hierarchy 相同的逻辑）
        if self.parallel_executor and merged_chunks:
            self.logger.info(f"[merge_by_content_length] 並行生成 {len(merged_chunks)} 個 call_prompts")
            prompts_list = []
            valid_indices = []
            for idx, chunk in enumerate(merged_chunks):
                prompt_dict = self._build_base_prompt_input(chunk)
                if prompt_dict:
                    prompts_list.append(prompt_dict)
                    valid_indices.append(idx)
                else:
                    # 内容過短，使用 fallback
                    merged_metadata = chunk.get('metadata', {})
                    heading = merged_metadata.get('heading', '')
                    struc_tags = merged_metadata.get('struc_tags', [])
                    contextual_tags = merged_metadata.get('contextual_tags', [])
                    content = chunk.get('content', '')
                    fallback_prompt = _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                    merged_metadata['call_prompt'] = fallback_prompt
            
            if prompts_list:
                # 並行调用
                results = self.parallel_executor.batch_chat(prompts_list)
                # 分配结果
                for (idx, _), result in zip([(valid_indices[i], prompts_list[i]) for i in range(len(valid_indices))], results):
                    merged_metadata = merged_chunks[idx].get('metadata', {})
                    if result:
                        # 清理响應
                        call_prompt = result.strip().strip('"\'').replace('\n', ' ').replace('\r', ' ')
                        # 移除 LLM 結束標記 <|eot_id|>
                        if '<|eot_id|>' in call_prompt:
                            call_prompt = call_prompt.split('<|eot_id|>')[0].strip()
                        if len(call_prompt) > 200:
                            call_prompt = call_prompt[:200].strip()
                        # 檢測錯誤消息
                        error_patterns = self.config.get('llm_error', {}).get('patterns', [])
                        if any(pattern in call_prompt for pattern in error_patterns):
                            # 使用 fallback
                            heading = merged_metadata.get('heading', '')
                            struc_tags = merged_metadata.get('struc_tags', [])
                            contextual_tags = merged_metadata.get('contextual_tags', [])
                            content = merged_chunks[idx].get('content', '')
                            call_prompt = _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                        
                        # 檢查是否需要過濾（長度 <= threshold）
                        if self.filter_casual_terms_enabled and len(call_prompt) <= self.filter_casual_terms_threshold:
                            if self._should_filter_single_term(call_prompt):
                                # 應該過濾，使用 fallback
                                heading = merged_metadata.get('heading', '')
                                struc_tags = merged_metadata.get('struc_tags', [])
                                contextual_tags = merged_metadata.get('contextual_tags', [])
                                content = merged_chunks[idx].get('content', '')
                                merged_metadata['call_prompt'] = _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                            else:
                                # 不需要過濾，保留原 call_prompt
                                merged_metadata['call_prompt'] = call_prompt
                        else:
                            merged_metadata['call_prompt'] = call_prompt
                    else:
                        # 使用 fallback
                        heading = merged_metadata.get('heading', '')
                        struc_tags = merged_metadata.get('struc_tags', [])
                        contextual_tags = merged_metadata.get('contextual_tags', [])
                        content = merged_chunks[idx].get('content', '')
                        merged_metadata['call_prompt'] = _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                
                # 並行判斷需要過濾的 call_prompts（已在上面處理，這裡不需要額外處理）
        else:
            # 降级：串行生成 call_prompts
            for chunk in merged_chunks:
                merged_metadata = chunk.get('metadata', {})
                if not merged_metadata.get('call_prompt'):
                    # 如果还沒有 call_prompt，生成它
                    heading = merged_metadata.get('heading', '')
                    os_tags = merged_metadata.get('os_tags', [])
                    struc_tags = merged_metadata.get('struc_tags', [])
                    contextual_tags = merged_metadata.get('contextual_tags', [])
                    content_tags = merged_metadata.get('content_tags', [])
                    content = chunk.get('content', '')
                    all_tags = os_tags + struc_tags + content_tags
                    unique_all_tags = list(dict.fromkeys(all_tags))
                    
                    call_prompt = _generate_base_prompt_with_llm(
                        content=content,
                        heading=heading,
                        tags=unique_all_tags,
                        os_tags=os_tags,
                        struc_tags=struc_tags,
                        contextual_tags=contextual_tags,
                        content_tags=content_tags,
                        llm_config=self.config.get('llm', {})
                    )
                    
                    # 檢查是否需要過濾（長度 <= threshold）
                    if self.filter_casual_terms_enabled and len(call_prompt) <= self.filter_casual_terms_threshold:
                        if self._should_filter_single_term(call_prompt):
                            # 應該過濾，使用 fallback
                            merged_metadata['call_prompt'] = _generate_fallback_base_prompt(heading, struc_tags, contextual_tags, content)
                        else:
                            # 不需要過濾，保留原 call_prompt
                            merged_metadata['call_prompt'] = call_prompt
                    else:
                        merged_metadata['call_prompt'] = call_prompt
        
        # 设置 ChunkIndex
        for idx, chunk in enumerate(merged_chunks):
            if isinstance(chunk, dict):
                chunk['ChunkIndex'] = idx + 1
                chunk['chunk_order'] = idx + 1
        
        return merged_chunks
        
    def _build_base_prompt_input(self, merged_chunk: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """
        构建 base_prompt 的输入字典（用於並行调用）
        
        Args:
            merged_chunk: 已合併的 chunk（包含 content 和 metadata）
            
        Returns:
            包含 prompt 信息的字典，如果内容過短则返回 None
        """
        merged_content = merged_chunk.get('content', '')
        metadata = merged_chunk.get('metadata', {})
        
        # 檢查内容是否足够
        content_clean = merged_content.strip() if merged_content else ''
        min_content_length = 10
        
        if len(content_clean) < min_content_length:
            return None
        
        # 准備输入数据
        heading = metadata.get('heading', '')
        os_tags = metadata.get('os_tags', [])
        struc_tags = metadata.get('struc_tags', [])
        contextual_tags = metadata.get('contextual_tags', [])
        content_tags = metadata.get('content_tags', [])
        all_tags = os_tags + struc_tags + content_tags
        
        # 加载 prompt 配置
        prompt_file = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'prompt', 'generate_base_prompt.json')
        if not os.path.exists(prompt_file):
            return None
        
        prompt_config = LOGger.load_json(prompt_file)
        if not prompt_config:
            return None
        
        # 构建 prompt
        system_prompt = prompt_config.get('system_prompt', '')
        user_prompt_template = prompt_config.get('user_prompt_template', '')
        max_tokens = int(prompt_config.get('generation_config', {}).get('max_new_tokens', 100))
        temperature = float(prompt_config.get('generation_config', {}).get('temperature', 0.3))
        
        # 格式化用户 prompt
        input_data = {
            'content': merged_content[:500] if len(merged_content) > 500 else merged_content,
            'heading': heading or '',
            'os_tags': os_tags or [],
            'struc_tags': struc_tags or [],
            'contextual_tags': contextual_tags or [],
            'content_tags': content_tags or [],
            'all_tags': all_tags or []
        }
        user_prompt = user_prompt_template.format(**input_data)
        
        # 獲取 LLM 配置
        llm_config = self.config.get('llm', {})
        chat_provider = llm_config.get('chat_provider', 'remote')
        chat_model = llm_config.get('chat_model', 'remote8b')
        
        # 模型名称转换
        if chat_provider == 'openai':
            if chat_model == 'gpt-3.5-turbo':
                chat_model = 'gpt35_chat'
            elif chat_model == 'gpt-4o':
                chat_model = 'gpt4o_chat'
            elif chat_model == 'gpt-4o-mini':
                chat_model = 'o4_chat'
            elif chat_model == 'gpt-4':
                chat_model = 'gpt4_chat'
        
        return {
            'prompt': user_prompt,
            'system_prompt': system_prompt,
            'provider': chat_provider,
            'model': chat_model,
            'max_tokens': max_tokens,
            'temperature': temperature
        }
    
    def _merge_chunk_group_content_only(self, group_chunks: List[Dict[str, Any]], group_index: int) -> Dict[str, Any]:
        """
        合併一组 chunks 的内容（不生成 call_prompt）
        
        Args:
            group_chunks: 要合併的 chunks 组
            group_index: 组索引（用於生成新的 ChunkIndex）
            
        Returns:
            合併后的 chunk（不包含 call_prompt）
        """
        return self._merge_chunk_group(group_chunks, group_index, skip_call_prompt=True)
    
    def _extract_original_prompts_from_chunk(self, chunk):
        """
        從單個 chunk 中提取原始的 call_prompt 和 multi_prompts
        
        Args:
            chunk: 單個 chunk 字典
            
        Returns:
            tuple: (original_call_prompt, original_multi_prompts)
        """
        metadata = chunk.get('metadata', {})
        
        # 檢查頂層和 metadata
        original_call_prompt = chunk.get('call_prompt') or metadata.get('call_prompt')
        original_multi_prompts = chunk.get('multi_prompts') or metadata.get('multi_prompts')
        
        return original_call_prompt, original_multi_prompts

    def _process_multi_prompts_for_segment(self, original_multi_prompts, use_multi_prompts, contextual_tags, content_tags, chunk_metadata=None):
        """
        處理 multi_prompts，優先使用原始值，否則基於標籤生成新的（用於 segment 構建）
        
        Args:
            original_multi_prompts: 原始的 multi_prompts
            use_multi_prompts: 是否啟用 multi_prompts 生成
            contextual_tags: 上下文標籤列表
            content_tags: 內容標籤列表
            chunk_metadata: chunk 的 metadata（用於提取 table_prompts）
            
        Returns:
            dict: 處理後的 multi_prompts 字典
        """
        multi_prompts = {}
        
        # 如果找到原始的 multi_prompts，優先使用它
        if original_multi_prompts:
            if isinstance(original_multi_prompts, dict):
                multi_prompts = original_multi_prompts.copy()
            elif isinstance(original_multi_prompts, list):
                # 列表格式轉換為字典格式（向後兼容）
                for idx, prompt_value in enumerate(original_multi_prompts):
                    if prompt_value and str(prompt_value).strip():
                        multi_prompts[f'multi_prompt_{idx}'] = str(prompt_value).strip()
            else:
                # 其他格式，嘗試轉換
                self.logger.warning(f"[_process_multi_prompts_segment] 原始 multi_prompts 格式不正確: {type(original_multi_prompts)}")
        elif use_multi_prompts:
            # 重新生成 multi_prompts（基於標籤）
            max_contextual_tags = 2  # 預設值
            
            # 步驟 1: 處理 contextual_tags
            filtered_contextual = self._filter_meaningless_tags(contextual_tags) if hasattr(self, '_filter_meaningless_tags') else contextual_tags
            filtered_contextual = [tag.strip() for tag in filtered_contextual if isinstance(tag, str) and tag.strip()]
            contextual_prompt_values = []
            if filtered_contextual and max_contextual_tags > 0:
                n = len(filtered_contextual)
                jump = max(1, math.ceil(n / max_contextual_tags)) if n > max_contextual_tags else 1
                for i in range(0, n, jump):
                    if i < n:
                        tags_slice = filtered_contextual[i:]
                        if tags_slice:
                            prompt_value = ' | '.join(tags_slice)
                            contextual_prompt_values.append(prompt_value)
                            if len(contextual_prompt_values) >= max_contextual_tags:
                                break
            
            # 步驟 2: 處理 content_tags
            content_prompt_values = []
            if content_tags:
                for tag in content_tags:
                    if not tag:
                        continue
                    tag_text = tag.strip() if isinstance(tag, str) else str(tag).strip()
                    if tag_text:
                        content_prompt_values.append(tag_text)
            
            # 步驟 2.5: 處理表格提示詞
            table_prompt_values = []
            if chunk_metadata:
                table_prompts = chunk_metadata.get('table_prompts', [])
                if table_prompts and isinstance(table_prompts, list):
                    for table_prompt in table_prompts:
                        if table_prompt and isinstance(table_prompt, str):
                            table_prompt_text = table_prompt.strip()
                            if table_prompt_text:
                                table_prompt_values.append(table_prompt_text)
            
            # 步驟 3: 組合生成 multi_prompts
            seen_prompts = set()
            if contextual_prompt_values:
                for ctx_idx, ctx_prompt in enumerate(contextual_prompt_values):
                    if ctx_prompt not in seen_prompts:
                        seen_prompts.add(ctx_prompt)
                        key = f'contextual_tags_{ctx_idx}'
                        multi_prompts[key] = ctx_prompt
            if content_prompt_values:
                for content_idx, content_prompt in enumerate(content_prompt_values):
                    if content_prompt not in seen_prompts:
                        seen_prompts.add(content_prompt)
                        key = f'content_tag_{content_idx}'
                        multi_prompts[key] = content_prompt
            if table_prompt_values:
                for table_idx, table_prompt in enumerate(table_prompt_values):
                    if table_prompt not in seen_prompts:
                        seen_prompts.add(table_prompt)
                        key = f'table_prompt_{table_idx}'
                        multi_prompts[key] = table_prompt
        
        return multi_prompts

    def _process_call_prompt_for_segment(self, original_call_prompt, content, max_length=200):
        """
        處理 call_prompt，優先使用原始值，否則截取內容（用於 segment 構建）
        
        Args:
            original_call_prompt: 原始的 call_prompt
            content: 內容文本
            max_length: 最大長度
            
        Returns:
            str: 處理後的 call_prompt
        """
        if original_call_prompt:
            return str(original_call_prompt).strip()
        
        # 使用內容截取作為 call_prompt
        if content:
            return content[:max_length] + "..." if len(content) > max_length else content
        
        return ""

    def _extract_original_prompts_from_chunks(self, chunks):
        """
        從 chunks 中提取原始的 call_prompt 和 multi_prompts
        
        Args:
            chunks: chunk 列表
            
        Returns:
            tuple: (original_call_prompt, original_multi_prompts)
        """
        original_call_prompt = None
        original_multi_prompts = None
        
        for i, chunk in enumerate(chunks, 1):
            # 檢查頂層
            if not original_call_prompt and chunk.get('call_prompt'):
                original_call_prompt = chunk['call_prompt']
                self.logger.debug(f"[_extract_original_prompts] 從 Chunk {i} 頂層找到 call_prompt")
            
            if not original_multi_prompts and chunk.get('multi_prompts'):
                original_multi_prompts = chunk['multi_prompts']
                self.logger.debug(f"[_extract_original_prompts] 從 Chunk {i} 頂層找到 multi_prompts")
            
            # 檢查 metadata
            if chunk.get('metadata'):
                if not original_call_prompt and chunk['metadata'].get('call_prompt'):
                    original_call_prompt = chunk['metadata']['call_prompt']
                    self.logger.debug(f"[_extract_original_prompts] 從 Chunk {i} metadata 找到 call_prompt")
                
                if not original_multi_prompts and chunk['metadata'].get('multi_prompts'):
                    original_multi_prompts = chunk['metadata']['multi_prompts']
                    self.logger.debug(f"[_extract_original_prompts] 從 Chunk {i} metadata 找到 multi_prompts")
            
            # 如果兩個都找到了，可以提前退出
            if original_call_prompt and original_multi_prompts:
                break
        
        return original_call_prompt, original_multi_prompts

    def _process_multi_prompts_for_merge(self, original_multi_prompts, unique_contextual_tags):
        """
        處理 multi_prompts，優先使用原始值，否則基於標籤生成新的
        
        Args:
            original_multi_prompts: 原始的 multi_prompts
            unique_contextual_tags: 用於生成的標籤列表
            
        Returns:
            dict: 處理後的 multi_prompts 字典
        """
        merged_multi_prompts = {}
        
        if original_multi_prompts:
            # 保留原始的 multi_prompts
            if isinstance(original_multi_prompts, dict):
                merged_multi_prompts = original_multi_prompts.copy()
            elif isinstance(original_multi_prompts, list):
                # 列表格式轉換為字典格式（向後兼容）
                for idx, prompt_value in enumerate(original_multi_prompts):
                    if prompt_value and str(prompt_value).strip():
                        merged_multi_prompts[f'multi_prompt_{idx}'] = str(prompt_value).strip()
            else:
                # 其他格式，嘗試轉換為字典
                self.logger.warning(f"[_process_multi_prompts] 原始 multi_prompts 格式不正確: {type(original_multi_prompts)}")
        else:
            # 重新生成 multi_prompts（基於合併後的標籤）
            # contextual_tags 已經由 [*os_tags, *struc_tags] 組成，直接使用
            filtered_contextual = self._filter_meaningless_tags(unique_contextual_tags) if hasattr(self, '_filter_meaningless_tags') else unique_contextual_tags
            filtered_contextual = [tag.strip() for tag in filtered_contextual if isinstance(tag, str) and tag.strip()]
            
            # 生成 multi_prompts（使用與 build_segments_from_chunks 相同的邏輯）
            max_contextual_tags = 2  # 預設值
            contextual_prompt_values = []
            if filtered_contextual and max_contextual_tags > 0:
                # 計算 jump（步長），確保最多生成 max_contextual_tags 個累加標籤
                n = len(filtered_contextual)
                jump = max(1, math.ceil(n / max_contextual_tags)) if n > max_contextual_tags else 1
                # 生成累加標籤：從不同位置開始的標籤組合
                for i in range(0, n, jump):
                    if i < n:
                        tags_slice = filtered_contextual[i:]
                        if tags_slice:
                            prompt_value = ' | '.join(tags_slice)
                            contextual_prompt_values.append(prompt_value)
                            # 如果已經達到 max_contextual_tags 個，停止
                            if len(contextual_prompt_values) >= max_contextual_tags:
                                break
            
            # 將 contextual_prompt_values 轉換為 multi_prompts 字典
            for idx, prompt_value in enumerate(contextual_prompt_values):
                if prompt_value:
                    key = f'contextual_tags_{idx}'
                    merged_multi_prompts[key] = prompt_value
        
        return merged_multi_prompts

    def _process_call_prompt_for_merge(self, original_call_prompt, merged_content, first_sentence):
        """
        處理 call_prompt，優先使用原始值，否則生成新的
        
        Args:
            original_call_prompt: 原始的 call_prompt
            merged_content: 合併後的內容
            first_sentence: 第一句話
            
        Returns:
            str: 處理後的 call_prompt
        """
        if original_call_prompt:
            # 使用原始的 call_prompt
            call_prompt_value = str(original_call_prompt).strip()
            self.logger.debug(f"[_process_call_prompt] 使用原始的 call_prompt: {call_prompt_value[:100]}...")
            return call_prompt_value
        
        # 重新生成 call_prompt（基於合併後的內容和標籤）
        call_prompt_value = None
        
        # 嘗試使用 LLM 生成 call_prompt
        if self.use_llm_for_call_prompt and merged_content:
            try:
                call_prompt_value = self._generate_call_prompt_with_llm(merged_content)
                self.logger.debug(f"[_process_call_prompt] 使用 LLM 生成 call_prompt")
            except Exception as e:
                self.logger.warning(f"[_process_call_prompt] LLM 生成 call_prompt 失敗: {e}")
        
        # 如果 LLM 生成失敗或不啟用，使用 fallback 方法
        if not call_prompt_value:
            # 使用第一句或內容摘要作為 call_prompt
            if first_sentence:
                call_prompt_value = first_sentence
            elif merged_content:
                # 截取前200字符作為 call_prompt
                call_prompt_value = merged_content[:200] + "..." if len(merged_content) > 200 else merged_content
            else:
                call_prompt_value = "合併的內容區塊"
        
        return call_prompt_value

    def _merge_chunk_group(self, group_chunks: List[Dict[str, Any]], group_index: int, skip_call_prompt: bool = False) -> Dict[str, Any]:
        """
        合併一組 chunks 成為單個 chunk
        
        Args:
            group_chunks: 要合併的 chunks 組
            group_index: 組索引（用於生成新的 ChunkIndex）
            
        Returns:
            合併後的 chunk
        """
        import uuid

        def _append_unique_normalized(target_list: List[str], values) -> None:
            """
            將來源標籤整理後加入列表並避免重複
            如果標籤包含 || 分隔符，會分割成多個獨立的標籤
            """
            if not values:
                return
            for value in values:
                if value is None:
                    continue
                if isinstance(value, str):
                    normalized = value.strip()
                else:
                    normalized = str(value).strip()
                
                if normalized:
                    # 如果標籤包含 || 分隔符，分割成多個獨立的標籤
                    if '||' in normalized:
                        split_tags = [t.strip() for t in normalized.split('||') if t.strip()]
                        for tag in split_tags:
                            if tag and tag not in target_list:
                                target_list.append(tag)
                    else:
                        if normalized not in target_list:
                            target_list.append(normalized)
        
        # 生成新的 ChunkID
        new_chunk_id = str(uuid.uuid4())
        
        aggregated_contextual_tags: List[str] = []
        aggregated_content_tags: List[str] = []
        
        for chunk in group_chunks:
            metadata = chunk.get('metadata', {}) or {}
            contextual_source = metadata.get('contextual_tags') or chunk.get('contextual_tags', [])
            content_source = metadata.get('content_tags') or chunk.get('content_tags', [])
            _append_unique_normalized(aggregated_contextual_tags, contextual_source)
            _append_unique_normalized(aggregated_content_tags, content_source)
        
        # 構建階層式內容
        merged_content = self._build_hierarchical_content(group_chunks)
        
        # 調試：檢查合併後的內容
        if not merged_content:
            self.logger.warning(f"[_merge_chunk_group] 合併後的內容為空，group_chunks 數量: {len(group_chunks)}")
            for i, chunk in enumerate(group_chunks):
                chunk_content = chunk.get('Content', '') or chunk.get('content', '')
                chunk_metadata = chunk.get('metadata', {})
                chunk_heading = chunk_metadata.get('heading', '')
                self.logger.warning(f"[_merge_chunk_group] Chunk {i}: Content='{chunk_content[:50] if chunk_content else '(空)'}', heading='{chunk_heading}'")
        
        # 從合併後的內容中提取第一句作為小標題
        # 第一句定義為：第一行，或第一個句號/換行符之前的內容（最多50字符）
        first_sentence = ''
        if merged_content:
            # 先嘗試取第一行
            first_line = merged_content.split('\n')[0].strip()
            if first_line:
                # 如果第一行太長，取第一個句號之前的部分
                if len(first_line) > 50:
                    # 找第一個句號、問號、驚嘆號或換行符
                    for sep in ['。', '.', '？', '?', '！', '!', '\n']:
                        idx = first_line.find(sep)
                        if idx > 0 and idx <= 50:
                            first_sentence = first_line[:idx].strip()
                            break
                    # 如果沒找到分隔符，直接截取前50字符
                    if not first_sentence:
                        first_sentence = first_line[:50].strip()
                else:
                    first_sentence = first_line
        
        # 從合併後的內容中提取 content_tags（使用 LLM 提取關鍵詞）
        unique_content_tags: List[str] = []
        if aggregated_content_tags:
            _append_unique_normalized(unique_content_tags, aggregated_content_tags)
        if merged_content:
            self.logger.debug(f"｜｜｜merged_content｜｜｜\n{merged_content[:200]}")
            # 使用 _extract_content_keywords 從合併後的內容提取關鍵詞作為 content_tags
            extracted_content_keywords = self._extract_content_keywords(merged_content)
            if extracted_content_keywords:
                _append_unique_normalized(unique_content_tags, extracted_content_keywords)
        
        # 收集 os_tags
        all_os_tags: List[str] = []
        for chunk in group_chunks:
            metadata = chunk.get('metadata', {}) or {}
            os_source = metadata.get('os_tags') or chunk.get('os_tags', [])
            _append_unique_normalized(all_os_tags, os_source)
        
        # 去重 os_tags
        unique_os_tags = list(dict.fromkeys(all_os_tags))  # 保持順序的去重
        
        # 收集 struc_tags：從 group_chunks 中收集所有 heading，按照 level 排序（階層序）
        # 階層序：level 小的在前（高階層在前）
        heading_level_pairs = []
        for chunk in group_chunks:
            metadata = chunk.get('metadata', {}) or {}
            level = metadata.get('level', chunk.get('level', 0))
            heading_value = metadata.get('heading') or chunk.get('heading') or chunk.get('title', '')
            heading = heading_value.strip() if isinstance(heading_value, str) else ''
            if heading:  # 只保留有值的 heading
                heading_level_pairs.append((level, heading))
        
        # 按照 level 排序（升序，高階層在前）
        heading_level_pairs.sort(key=lambda x: x[0])
        
        # 提取 struc_tags（按照階層序）
        unique_struc_tags = [heading for level, heading in heading_level_pairs]
        # 去重但保持順序
        unique_struc_tags = list(dict.fromkeys(unique_struc_tags))
        
        # contextual_tags 依序由 [*os_tags, *struc_tags] 組成，並保留原 chunk 的脈絡標籤
        unique_contextual_tags = list(aggregated_contextual_tags)
        for tag in unique_os_tags + unique_struc_tags:
            if tag not in unique_contextual_tags:
                unique_contextual_tags.append(tag)
        
        # 找出最高階層的標題作為 heading（用於 metadata）
        min_level = float('inf')
        root_heading_for_struc = ''
        for chunk in group_chunks:
            metadata = chunk.get('metadata', {})
            level = metadata.get('level', 0)
            heading = metadata.get('heading', '').strip()
            # 找出最高階層的標題（level 最小且非空）
            if heading and level < min_level:
                min_level = level
                root_heading_for_struc = heading
        
        # 如果沒有找到最高階層標題，使用第一句作為 heading
        if not root_heading_for_struc and first_sentence:
            root_heading_for_struc = first_sentence
        
        # 提取原始的 call_prompt 和 multi_prompts
        original_call_prompt, original_multi_prompts = self._extract_original_prompts_from_chunks(group_chunks)
        
        # 處理 multi_prompts
        merged_multi_prompts = self._process_multi_prompts_for_merge(original_multi_prompts, unique_contextual_tags)
        
        # 處理 call_prompt
        if skip_call_prompt:
            call_prompt = ''
        else:
            call_prompt = self._process_call_prompt_for_merge(original_call_prompt, merged_content, first_sentence)
            
            # 如果沒有原始 call_prompt 且需要使用 LLM 生成更詳細的 call_prompt
            if not original_call_prompt and self.use_llm_for_call_prompt:
                heading_for_prompt = first_sentence or root_heading_for_struc or ''
                all_tags = unique_os_tags + unique_struc_tags + unique_content_tags
                unique_all_tags = list(dict.fromkeys(all_tags))
                
                try:
                    llm_call_prompt = _generate_base_prompt_with_llm(
                        content=merged_content,
                        heading=heading_for_prompt,
                        tags=unique_all_tags,
                        os_tags=unique_os_tags,
                        struc_tags=unique_struc_tags,
                        contextual_tags=unique_contextual_tags,
                        content_tags=unique_content_tags,
                        llm_config=getattr(self, 'config', {}).get('llm', {})
                    )
                    if llm_call_prompt:
                        call_prompt = llm_call_prompt
                except Exception as e:
                    self.logger.warning(f"[_merge_chunk_group] LLM 生成詳細 call_prompt 失敗: {e}")
                    # 保持使用 _process_call_prompt_for_merge 的結果
        
        # 構建合併後的 metadata
        merged_metadata = group_chunks[0].get('metadata', {}).copy()
        # 更新 heading 為第一句（小標題）
        if first_sentence:
            merged_metadata['heading'] = first_sentence
        elif root_heading_for_struc:
            merged_metadata['heading'] = root_heading_for_struc
        # 更新 level 為最高階層（最小 level）
        if min_level != float('inf'):
            merged_metadata['level'] = min_level
        
        merged_metadata.update({
            'os_tags': unique_os_tags,
            'struc_tags': unique_struc_tags,
            'contextual_tags': unique_contextual_tags,  # 合併版本，向後兼容
            'content_tags': unique_content_tags,
            'call_prompt': call_prompt,  # 合併後的 call_prompt
            'multi_prompts': merged_multi_prompts,  # 合併後的 multi_prompts
            'is_merged': True,
            'chunks_count': len(group_chunks)
        })
        
        # 使用第一個 chunk 的基本屬性作為模板
        base_chunk = group_chunks[0]
        
        # 從所有 chunks 中找出 source（優先使用頂層 source，否則從 metadata 取得）
        merged_source = None
        self.logger.info(f"[_merge_chunk_group] 開始從 {len(group_chunks)} 個 chunks 中查找 source")
        for i, chunk in enumerate(group_chunks):
            if isinstance(chunk, dict):
                source = chunk.get('source')
                self.logger.debug(f"[_merge_chunk_group] Chunk {i+1} 頂層 source: {source}")
                if not source:
                    metadata = chunk.get('metadata', {})
                    source = metadata.get('document_source') or metadata.get('file_path')
                    self.logger.debug(f"[_merge_chunk_group] Chunk {i+1} metadata source: {source}")
                if source:
                    merged_source = source
                    break
        
        if merged_source:
            merged_metadata['document_source'] = merged_source
            self.logger.info(f"[_merge_chunk_group] 合併後的 chunk 使用 source: {merged_source}")
        else:
            self.logger.warning(f"[_merge_chunk_group] 無法從 {len(group_chunks)} 個 chunks 中找到 source，將使用空字符串")
            # 詳細記錄每個 chunk 的結構以便調試
            for i, chunk in enumerate(group_chunks):
                self.logger.warning(f"[_merge_chunk_group] Chunk {i+1} 結構: keys={list(chunk.keys()) if isinstance(chunk, dict) else 'not a dict'}")
                if isinstance(chunk, dict) and 'metadata' in chunk:
                    self.logger.warning(f"[_merge_chunk_group] Chunk {i+1} metadata keys: {list(chunk['metadata'].keys())}")
        
        # group_index 是 0-based，ChunkIndex 應該是 1-based（在文件中的順序）
        final_chunk_index = group_index + 1
        
        merged_chunk = {
            'ChunkID': new_chunk_id,
            'content': merged_content,
            'DocumentIndexID': base_chunk.get('DocumentIndexID'),
            'ChunkIndex': final_chunk_index,  # 組合後在文件中的順序（從 1 開始）
            'Page': base_chunk.get('Page'),
            'Language': base_chunk.get('Language', 'default'),
            'RegisteredGroup': base_chunk.get('RegisteredGroup', ''),
            'CreateUser': base_chunk.get('CreateUser', 'system'),
            'IsActivate': True,
            'IsDelete': False,
            'source': merged_source or '',  # 確保有 source 欄位（即使為空）
            'metadata': merged_metadata,
            # 將 call_prompt 和 multi_prompts 也存儲在 chunk 的頂層（向後兼容）
            'call_prompt': call_prompt,
            'multi_prompts': merged_multi_prompts,
            'chunk_order': final_chunk_index,  # 與 ChunkIndex 一致
            'os_tags': unique_os_tags,
            'struc_tags': unique_struc_tags,
            'contextual_tags': unique_contextual_tags,
            'content_tags': unique_content_tags
        }
        
        self.logger.debug(f"[_merge_chunk_group] 合併 {len(group_chunks)} 個 chunks，設置 ChunkIndex={final_chunk_index}")
        return merged_chunk
    
    def _build_hierarchical_content(self, group_chunks: List[Dict[str, Any]]) -> str:
        """
        按階層組織內容，區分標題和內文，避免標題被當成內文合併
        
        Args:
            group_chunks: 要合併的 chunks 組
            
        Returns:
            階層式組織的內容字符串（標題和內文分開處理）
        """
        if not group_chunks:
            return ''
        
        def _is_heading_chunk(chunk):
            """判斷 chunk 是否是標題"""
            metadata = chunk.get('metadata', {})
            heading = metadata.get('heading', '') or chunk.get('heading', '') or chunk.get('title', '')
            if not heading or not heading.strip():
                return False
            
            content = chunk.get('Content', '') or chunk.get('content', '')
            if not content:
                content = metadata.get('content', '') or metadata.get('text', '')
            
            if content and content.strip():
                content_stripped = content.strip()
                heading_stripped = heading.strip()
                # 如果 content 等於 heading，或者 content 很短且包含 heading
                if content_stripped == heading_stripped or (len(content_stripped) < 100 and heading_stripped in content_stripped):
                    return True
            else:
                return True
            
            return False
        
        # 找出最高階層（level 最小的）的標題
        min_level = float('inf')
        root_heading = ''
        heading_chunks = []  # 標題 chunks（按 level 排序）
        content_chunks = []  # 內文 chunks
        
        for idx, chunk in enumerate(group_chunks):
            metadata = chunk.get('metadata', {})
            level = metadata.get('level', 0)
            heading = metadata.get('heading', '').strip()
            
            # 判斷是否是標題
            is_heading = _is_heading_chunk(chunk)
            
            if is_heading:
                # 標題 chunk，記錄層級信息並加入 merged_content
                heading_chunks.append((level, heading, chunk))
                # 找出最高階層的標題
                if heading and level < min_level:
                    min_level = level
                    root_heading = heading
                
                # 標題 chunk 的 content 也要加入 merged_content（這樣才能保留層級結構）
                content = chunk.get('Content', '') or chunk.get('content', '')
                if not content:
                    content = metadata.get('content', '') or metadata.get('text', '')
                if isinstance(content, str):
                    content = content.strip()
                else:
                    content = str(content).strip() if content else ''
                
                if content:
                    content_chunks.append(content)
                    self.logger.debug(f"[_build_hierarchical_content] 包含標題內容: level={level}, content='{content[:50]}...'")
            else:
                # 內文 chunk，加入 merged_content
                content = chunk.get('Content', '') or chunk.get('content', '')
                if not content:
                    content = metadata.get('content', '') or metadata.get('text', '')
                if isinstance(content, str):
                    content = content.strip()
                else:
                    content = str(content).strip() if content else ''
                
                if content:
                    content_chunks.append(content)
        
        # 構建合併後的內容
        # 注意：標題 chunk 的 content 不加入 merged_content，避免標題被當成內文
        # 標題只用於 root_heading（在 metadata 中），不加入 merged_content
        merged_content_parts = []
        
        # 只合併內文 chunks，不加入標題文字
        if content_chunks:
            merged_content = '\n\n'.join(content_chunks)
            if merged_content:
                merged_content_parts.append(merged_content)
        
        result = '\n\n'.join(merged_content_parts)
        
        # 如果結果為空，至少返回第一個 chunk 的 Content 或 heading
        if not result:
            first_chunk = group_chunks[0] if group_chunks else {}
            first_content = first_chunk.get('Content', '') or first_chunk.get('content', '')
            if isinstance(first_content, str):
                first_content = first_content.strip()
            else:
                first_content = str(first_content).strip() if first_content else ''
            
            if not first_content:
                first_metadata = first_chunk.get('metadata', {})
                first_heading = first_metadata.get('heading', '').strip()
                if first_heading:
                    result = first_heading
        
        return result
    
    def _extract_group_heading_from_buffer(self, buffer_groups: List[Dict[str, Any]], chunks: List[Dict[str, Any]], 
                                          _get_chunk_level, _get_chunk_heading) -> str:
        """
        從 buffer 中的所有群組中，按照書寫順序提取 level=0 標題
        
        Args:
            buffer_groups: buffer 中的群組列表，每個群組包含 'indices' 和 'heading'
            chunks: 原始 chunks 列表
            _get_chunk_level: 獲取 chunk level 的函數
            _get_chunk_heading: 獲取 chunk heading 的函數
            
        Returns:
            合併後的標題字符串（多個 level=0 標題用 \n 連接），如果沒有找到則返回第一個群組的標題
        """
        # 收集 buffer 中所有群組的所有 chunk 索引
        all_indices = []
        for group in buffer_groups:
            all_indices.extend(group['indices'])
        
        # 去重並排序（按照書寫順序）
        all_indices = sorted(set(all_indices))
        
        # 按照書寫順序找出所有 level=0 且有 heading 的 chunks
        level0_headings = []
        for idx in all_indices:
            if idx < 0 or idx >= len(chunks):
                continue
            chunk = chunks[idx]
            level = _get_chunk_level(chunk)
            heading = _get_chunk_heading(chunk)
            
            # 如果是 level=0 且有標題，記錄下來
            if level == 0 and heading and heading.strip():
                heading_clean = heading.strip()
                # 避免重複
                if heading_clean not in level0_headings:
                    level0_headings.append(heading_clean)
        
        # 如果找到多個 level=0 標題，用 \n 合併
        if level0_headings:
            return '\n'.join(level0_headings)
        
        # Fallback: 使用 buffer 中第一個群組的標題
        if buffer_groups and buffer_groups[0].get('heading'):
            return buffer_groups[0]['heading']
        
        return ''
    
    def _extract_root_level_tags(self, group_chunks: List[Dict[str, Any]]) -> List[str]:
        """
        提取回溯到文件最上層的標籤
        
        Args:
            group_chunks: chunks 組
            
        Returns:
            根層級標籤列表
        """
        root_tags = []
        
        # 找出所有 chunks 中最高層級的標題
        min_level = float('inf')
        root_headings = []
        
        for chunk in group_chunks:
            metadata = chunk.get('metadata', {})
            level = metadata.get('level', 0)
            heading = metadata.get('heading', '')
            
            if level < min_level and heading and heading.strip():
                min_level = level
                root_headings = [heading.strip()]
            elif level == min_level and heading and heading.strip():
                if heading.strip() not in root_headings:
                    root_headings.append(heading.strip())
        
        # 如果找到根層級標題，添加到標籤中
        if root_headings and min_level != float('inf'):
            root_tags.extend(root_headings)
        
        return root_tags
    
    def _segment_by_lines(self, text: str) -> List[str]:
        """按行分段"""
        return segment_by_lines(text)
    
    def _segment_by_paragraphs(self, text: str) -> List[str]:
        """按段落分段"""
        return segment_by_paragraphs(text)
    
    def _segment_with_textprocessor(self, text: str, filename: str = "") -> List[str]:
        """使用 TextProcessor 進行語義分段"""
        try:
            # 獲取 TextProcessor 配置
            textprocessor_config = self.config.get('llm', {})
            base_url = textprocessor_config.get('base_url', 'http://10.1.3.127:6017')
            
            # 準備請求
            payload = {
                "text": text,
                "llm_config": {
                    "provider": textprocessor_config.get('chat_provider', 'remote'),
                    "model": textprocessor_config.get('chat_model', 'remote8b'),
                    "max_new_tokens": 500,
                    "temperature": 0.3
                },
                "filename": filename,
                "include_filename": bool(filename),
                "use_multi_prompts": True,
                "segment_mode": "content"  # 返回完整內容
            }
            
            # 發送請求
            response = requests.post(
                f"{base_url}/segment/segment-text-with-llm",
                json=payload,
                timeout=textprocessor_config.get('timeout', 60)
            )
            
            if response.status_code == 200:
                result = response.json()
                segments = result.get('segments', [])
                
                # 提取分段內容（檢查多個可能的字段）
                chunks = []
                total_content_length = 0
                for segment in segments:
                    # 優先使用 content，如果沒有則嘗試其他字段
                    content = segment.get('content', '') or segment.get('text', '') or segment.get('full_content', '')
                    content = content.strip() if content else ''
                    if content:
                        chunks.append(content)
                        total_content_length += len(content)
                
                # 驗證：如果返回的內容總長度遠小於輸入（少於 30%），可能內容不完整
                if chunks:
                    coverage_ratio = total_content_length / len(text) if text else 0
                    self.logger.info(f"TextProcessor 分段成功: {len(chunks)} 個片段，內容覆蓋率: {coverage_ratio:.2%}")
                    
                    # 如果覆蓋率太低，回退到段落分段
                    if coverage_ratio < 0.3:
                        self.logger.warning(f"TextProcessor 返回內容覆蓋率過低 ({coverage_ratio:.2%})，回退到段落分段")
                        return self._segment_by_paragraphs(text)
                    
                    return chunks
                else:
                    self.logger.warning("TextProcessor 返回空分段結果")
            else:
                self.logger.warning(f"TextProcessor 請求失敗: {response.status_code}")
                
        except Exception as e:
            self.logger.warning(f"TextProcessor 分段失敗: {e}")
        
        # 失敗時回退到段落分段
        return self._segment_by_paragraphs(text)
    
    def _segment_by_custom_pattern(self, text: str, sep: str) -> List[str]:
        """根據自定義模式分段"""
        return segment_by_custom_pattern(text, sep)

    def _analyze_text_structure(self, text: str) -> dict:
        """
        分析文本結構特徵
        
        Args:
            text: 輸入文字
            
        Returns:
            結構分析結果
        """
        import re
        
        # 載入分析配置
        try:
            prompt_config = self._load_prompt_config('text_structure_analysis.json')
            patterns = prompt_config.get('patterns', {})
        except:
            patterns = {}
        
        # 快速模式識別（不使用LLM）
        lines = text.split('\n')
        non_empty_lines = [line.strip() for line in lines if line.strip()]
        
        if not non_empty_lines:
            return {'recommended_strategy': 'paragraph'}
        
        # 檢測聯絡人列表格式 [公司] 姓名
        contact_pattern = re.compile(r'^\[.+?\].+')
        contact_matches = sum(1 for line in non_empty_lines if contact_pattern.match(line))
        
        if contact_matches / len(non_empty_lines) > 0.7:
            return {
                'structure_type': '聯絡人列表',
                'recommended_strategy': 'line',
                'reasoning': '檢測到大量 [公司] 姓名 格式，適合按行分段'
            }
        
        # 檢測配置文件格式
        config_pattern = re.compile(r'^[\w_]+\s*[=:]|^\[.+\]$')
        config_matches = sum(1 for line in non_empty_lines if config_pattern.match(line))
        
        if config_matches / len(non_empty_lines) > 0.5:
            return {
                'structure_type': '配置文件',
                'recommended_strategy': 'line',
                'reasoning': '檢測到鍵值對或節段格式，適合按行分段'
            }
        
        # 檢測 Markdown 表格（優先於其他結構檢測）
        # 表格特徵：包含 | 分隔符的行，且有分隔行（如 |---|---|）
        table_row_pattern = re.compile(r'^\|.*\|.*$')  # 表格行：以 | 開頭和結尾
        table_separator_pattern = re.compile(r'^\|[\s\-\|:]+\|$')  # 表格分隔行：|---|---| 或 |:---|:---:|---:|
        
        table_rows = sum(1 for line in non_empty_lines if table_row_pattern.match(line.strip()))
        table_separators = sum(1 for line in non_empty_lines if table_separator_pattern.match(line.strip()))
        
        # 如果表格行佔比超過 30%，且至少有 1 個分隔行，視為表格文檔
        if table_rows > 0 and table_separators > 0:
            table_ratio = table_rows / len(non_empty_lines) if non_empty_lines else 0
            if table_ratio > 0.3:
                return {
                    'structure_type': 'Markdown 表格',
                    'recommended_strategy': 'paragraph',
                    'reasoning': f'檢測到 Markdown 表格結構（表格行佔比 {table_ratio:.1%}），使用段落分段以保持每行記錄的完整性'
                }
        
        # 檢測 Markdown 或結構化文檔（包含標題、列表等）
        markdown_pattern = re.compile(r'^#{1,6}\s+|^[\*\-]\s+|^\d+\.\s+|^##\s+')
        markdown_matches = sum(1 for line in non_empty_lines if markdown_pattern.match(line))
        
        # 檢測結構化文檔特徵（標題層次、列表項目）
        has_titles = bool(re.search(r'^#{1,6}\s+', text, re.MULTILINE))
        has_lists = bool(re.search(r'^[\*\-]\s+|^\d+\.\s+', text, re.MULTILINE))
        has_paragraphs = text.count('\n\n') > 0
        
        # 如果同時包含標題、列表和段落，視為結構化文檔，使用語義分段
        # 但如果表格行佔比也很高，優先使用段落分段
        if (has_titles or has_lists) and (markdown_matches / len(non_empty_lines) > 0.1 or has_paragraphs):
            # 如果表格行也很多，使用段落分段而不是語義分段
            if table_rows > 0:
                table_ratio = table_rows / len(non_empty_lines) if non_empty_lines else 0
                if table_ratio > 0.2:  # 表格行佔比超過 20%
                    return {
                        'structure_type': '混合結構（標題+表格）',
                        'recommended_strategy': 'paragraph',
                        'reasoning': f'檢測到 Markdown 標題和表格混合結構（表格行佔比 {table_ratio:.1%}），使用段落分段以保持表格記錄完整性'
                    }
            
            return {
                'structure_type': '結構化文檔',
                'recommended_strategy': 'semantic',
                'reasoning': '檢測到 Markdown 標題或列表結構，使用語義分段以保持語義完整性'
            }
        
        # 檢測短行列表（平均行長度較短）
        avg_line_length = sum(len(line) for line in non_empty_lines) / len(non_empty_lines)
        
        if avg_line_length < 50 and len(non_empty_lines) > 10:
            return {
                'structure_type': '短條目列表',
                'recommended_strategy': 'line',
                'reasoning': f'平均行長度 {avg_line_length:.1f} 字符，適合按行分段'
            }
        
        # 檢測段落結構（有空行分隔）
        empty_line_count = text.count('\n\n')
        if empty_line_count > 2:
            return {
                'structure_type': '段落文檔',
                'recommended_strategy': 'paragraph',
                'reasoning': f'檢測到 {empty_line_count} 個段落分隔，適合按段落分段'
            }
        
        # 預設策略
        return {
            'structure_type': '一般文本',
            'recommended_strategy': 'paragraph',
            'reasoning': '未檢測到特殊結構，使用預設段落分段'
        }

    def _intelligent_text_segmentation(self, text: str, filename: str = "", custom_separator: str = None) -> List[str]:
        """
        智能文本分段 - 根據文本結構選擇最佳分段策略
        
        Args:
            text: 輸入文字
            filename: 檔案名稱（用於分析提示）
            custom_separator: 自定義分隔符（如果提供，優先使用此分隔符）
            
        Returns:
            分段後的文字列表
        """
        try:
            # 如果提供了自定義分隔符，優先使用
            if custom_separator is not None:
                self.logger.info(f"使用自定義分隔符進行分段: {repr(custom_separator)}")
                return self._segment_by_custom_pattern(text, custom_separator)
            
            # 1. 快速結構檢測
            structure_info = self._analyze_text_structure(text)
            
            # 2. 根據結構選擇分段策略
            strategy = structure_info.get('recommended_strategy', 'paragraph')
            
            self.logger.info(f"文本結構分析: {structure_info.get('structure_type', '未知')}, 採用策略: {strategy}")
            
            if strategy == 'line':
                return self._segment_by_lines(text)
            elif strategy == 'paragraph':
                return self._segment_by_paragraphs(text)
            elif strategy == 'semantic':
                return self._segment_with_textprocessor(text, filename)
            elif strategy == 'custom':
                sep = structure_info.get('custom_pattern', '\n')
                return self._segment_by_custom_pattern(text, sep)
            else:
                # 預設使用段落分段
                return self._segment_by_paragraphs(text)
                
        except Exception as e:
            LOGger.exception_process(e, logfile='', stamps=['ContextParser','_intelligent_text_segmentation'])
            self.logger.error(f"智能分段失敗，使用改進的簡單分段: {e}")
            return self._split_text(text, max_length=1000, min_length=50)
    

def batch_process(method, inputs: List[Dict[str, Any]], batch_size: int, workers: int, params: Optional[Dict[str, Any]] = None, max_workers: int = 4, max_batch_size: int = 100, **kwargs) -> Dict[str, Any]:
    """用多執行緒池批量處理文件，並返回處理結果"""
    params = params or {}
    # 限制參數範圍，避免資源過度消耗
    batch_size = min(max(1, batch_size), max_batch_size)
    workers = min(max(1, workers), max_workers)
    
    params = params or {}
    all_results = []
    
    # 分批處理
    for i in range(0, len(inputs), batch_size):
        batch = inputs[i:i + batch_size]
        
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = [
                executor.submit(
                    method,
                    *args,
                    **params
                ) for args in batch
            ]
            batch_results = [future.result() for future in futures]
            all_results.extend(batch_results)
    
    return all_results

class DocumentParser(ABC):
    """抽象類別"""
    def __init__(self, 
            config: Dict[str, Any] = None, 
            llm_service=None, 
            structure_mode: str = "hybrid", 
            separator: str = '\n\n',
            **kwargs
        ):
        self.config = config
        self.llm_service = llm_service
        self.structure_mode = structure_mode
        self.separator = separator
        self.logger = m_logger

    @abstractmethod
    def open(self, **kwargs) -> bool:
        """開啟文件"""
        pass

    @abstractmethod
    def preview(self, **kwargs) -> Dict[str, Any]:
        """預覽文件"""
        pass
    
    @abstractmethod
    def parse(self, **kwargs) -> Dict[str, Any]:
        """初步解析文件"""
        pass

    @abstractmethod
    def extract(self, **kwargs) -> Dict[str, Any]:
        """將文件相同格式的文字分到最小單位，並賦予原有書寫時的 level 以及 書寫順序 order 等資訊"""
        pass

    @abstractmethod
    def segment(self, **kwargs) -> Dict[str, Any]:
        """參考 level 或 order 或字數或文意，將文件分段"""
        pass

    @abstractmethod
    def chunk(self, **kwargs) -> Dict[str, Any]:
        """為分段後的文字賦予 chunk_id 以及 chunk_order 等資訊"""
        pass

    def extract_batch(self, inputs: List[Dict[str, Any]], batch_size: int = 10, workers: int = 4, params: Optional[Dict[str, Any]] = None, max_workers: int = 4, max_batch_size: int = 100, **kwargs) -> Dict[str, Any]:
        """用多執行緒池批量處理文件，並返回處理結果"""
        all_results = batch_process(self.extract, inputs, batch_size, workers, params, max_workers, max_batch_size, **kwargs)

    def segment_batch(self, inputs: List[Dict[str, Any]], batch_size: int = 10, workers: int = 4, params: Optional[Dict[str, Any]] = None, max_workers: int = 4, max_batch_size: int = 100, **kwargs) -> Dict[str, Any]:
        """用多執行緒池批量處理文件，並返回處理結果"""
        all_results = batch_process(self.segment, inputs, batch_size, workers, params, max_workers, max_batch_size, **kwargs)

    def chunk_batch(self, inputs: List[Dict[str, Any]], batch_size: int = 10, workers: int = 4, params: Optional[Dict[str, Any]] = None, max_workers: int = 4, max_batch_size: int = 100, **kwargs) -> Dict[str, Any]:
        """用多執行緒池批量處理文件，並返回處理結果"""
        all_results = batch_process(self.chunk, inputs, batch_size, workers, params, max_workers, max_batch_size, **kwargs)

    def parse_batch(self, inputs: List[Dict[str, Any]], batch_size: int = 10, workers: int = 4, params: Optional[Dict[str, Any]] = None, max_workers: int = 4, max_batch_size: int = 100, **kwargs) -> Dict[str, Any]:
        """用多執行緒池批量處理文件，並返回處理結果"""
        all_results = batch_process(self.parse, inputs, batch_size, workers, params, max_workers, max_batch_size, **kwargs)
