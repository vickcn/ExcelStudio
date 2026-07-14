#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Standalone helpers for xlsx_parser (no context_parser / package / src deps)."""
from __future__ import annotations

import base64
import hashlib
import io
import json
import logging
import os
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

try:
    from PIL import Image
    HAS_PILLOW = True
except Exception:
    Image = None
    HAS_PILLOW = False

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_BATCH_SIZE_LIMIT = 190



def load_config(config: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    if isinstance(config, dict) and config:
        return config
    cfg_path = os.path.join(PROJECT_ROOT, "config.json")
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            pass
    return {}


def _project_prompt_path(name: str) -> Path:
    cfg = load_config()
    prompt_path = cfg.get("prompt_path")
    if prompt_path:
        return Path(prompt_path) / name
    return Path(PROJECT_ROOT) / "prompt" / name

IMAGE_EXPLAINER_FAILURE_LINE = "[圖像說明失敗]"


def _looks_like_refusal_response(text: str) -> bool:
    """辨識模型拒答／無法協助類訊息，避免當成圖像描述。"""
    if text is None:
        return True
    s = str(text).strip()
    if not s:
        return True
    if s == IMAGE_EXPLAINER_FAILURE_LINE:
        return True
    low = s.lower()
    needles_zh = (
        "無法協助",
        "無法滿足",
        "無法處理該要求",
        "不能協助",
        "抱歉，我無法",
        "對不起，我無法",
    )
    needles_en = (
        "cannot assist",
        "can't assist",
        "can not assist",
        "unable to assist",
        "unable to comply",
        "unable to fulfill",
        "i'm unable to",
        "i cannot assist",
        "i can't assist",
        "cannot help with",
    )
    for n in needles_zh:
        if n in s:
            return True
    for n in needles_en:
        if n in low:
            return True
    if "抱歉" in s and ("無法" in s or "不能" in s):
        return True
    if "sorry" in low and ("can't" in low or "cannot" in low or "unable" in low):
        return True
    return False


def _to_bool(value: Any) -> Optional[bool]:
    if isinstance(value, bool):
        return value
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().lower()
    if text in ("true", "1", "yes", "y", "是", "有"):
        return True
    if text in ("false", "0", "no", "n", "否", "無"):
        return False
    return None


def _extract_markdown_table(text: str) -> Optional[str]:
    if not text:
        return None
    lines = text.splitlines()
    n = len(lines)
    i = 0
    while i < n:
        if "|" not in lines[i]:
            i += 1
            continue
        j = i
        block = []
        while j < n and "|" in lines[j]:
            block.append(lines[j].rstrip())
            j += 1
        if len(block) >= 2 and any(_TABLE_SEPARATOR_RE.match(row.strip()) for row in block):
            return "\n".join(row.strip() for row in block if row.strip())
        i = j
    return None


def _extract_json_block(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None
    code_block = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL | re.IGNORECASE)
    if code_block:
        try:
            parsed = json.loads(code_block.group(1))
            if isinstance(parsed, dict):
                return parsed
        except Exception:
            pass
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        try:
            parsed = json.loads(text[start : end + 1])
            if isinstance(parsed, dict):
                return parsed
        except Exception:
            pass
    return None


def _remove_table_block(text: str, table_block: Optional[str]) -> str:
    if not text:
        return ""
    if not table_block:
        return text.strip()
    cleaned = text.replace(table_block, "")
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


def parse_image_explainer_raw_output(raw_output: str) -> Dict[str, Any]:
    """
    將 LLM 原始字串解析為結構化欄位（例如 has_table、summary、markdown_table）。
    支援 JSON 輸出與純文字回覆，並嘗試擷取 markdown 表格。
    """
    if raw_output is None:
        text = ""
    else:
        text = raw_output if isinstance(raw_output, str) else str(raw_output)
    json_obj = _extract_json_block(text)
    has_table: Optional[bool] = None
    summary: Optional[str] = None
    markdown_table: Optional[str] = None

    if isinstance(json_obj, dict):
        has_table = _to_bool(json_obj.get("has_table"))
        if has_table is None:
            has_table = _to_bool(json_obj.get("is_table"))
        md_value = json_obj.get("markdown_table")
        if md_value is None:
            md_value = json_obj.get("table")
        if md_value is not None:
            markdown_table = str(md_value).strip() or None
        summary_value = json_obj.get("summary")
        if summary_value is None:
            summary_value = json_obj.get("text")
        if summary_value is not None:
            summary = str(summary_value).strip() or None
    else:
        low = text.lower()
        if '"has_table"' in low:
            if re.search(r'"has_table"\s*:\s*true', low):
                has_table = True
            elif re.search(r'"has_table"\s*:\s*false', low):
                has_table = False
        md_like = _extract_json_like_field(text, "markdown_table")
        if md_like:
            markdown_table = md_like
        summary_like = _extract_json_like_field(text, "summary")
        if summary_like:
            summary = summary_like

    if not markdown_table:
        markdown_table = _extract_markdown_table(text)

    if has_table is None:
        if markdown_table:
            has_table = True
        elif re.search(r"(沒有|無)\s*表格", text):
            has_table = False
        elif "表格" in text:
            has_table = True

    if not summary:
        summary_match = re.search(r"(?:總結|摘要)\s*[：:]\s*(.+)$", text, re.DOTALL)
        if summary_match:
            summary = summary_match.group(1).strip()
        else:
            summary = _remove_table_block(text, markdown_table)

    return {
        "has_table": has_table,
        "summary": summary or "",
        "markdown_table": markdown_table,
        "raw": text,
    }


def image_explainer_table_output_valid(parsed: Dict[str, Any]) -> bool:
    """
    當 has_table 為 True 時，檢查輸出是否含約定之 markdown 表格（如前 10x10）。
    """
    if not isinstance(parsed, dict):
        return False
    if parsed.get("has_table") is not True:
        return True
    table = parsed.get("markdown_table")
    if table is None:
        table = _extract_markdown_table(str(parsed.get("raw", "")))
    if not table:
        return False
    lines = [line.strip() for line in str(table).splitlines() if line.strip()]
    if len(lines) < 2:
        return False
    if not any(_TABLE_SEPARATOR_RE.match(line) for line in lines):
        return False
    return True


def finalize_image_explainer_text(
    raw_output: str,
    *,
    max_format_retries: int = 0,
) -> str:
    """
    串接 parse ->（has_table 時）validate -> 不符則重試或降級 -> 回傳最終寫入 summary 的文字。
    目前先實作降級策略；max_format_retries 參數保留給後續重試機制。
    """
    _ = max_format_retries
    if raw_output is None:
        return ""
    parsed = parse_image_explainer_raw_output(raw_output if isinstance(raw_output, str) else str(raw_output))
    if parsed.get("has_table") is True:
        if image_explainer_table_output_valid(parsed):
            table_text = parsed.get("markdown_table")
            if table_text:
                out = str(table_text).strip()
                if _looks_like_refusal_response(out):
                    return IMAGE_EXPLAINER_FAILURE_LINE
                return out
        # 降級：有表但格式不符時，回傳摘要避免落空
        fallback_summary = parsed.get("summary")
        if fallback_summary:
            out = str(fallback_summary).strip()
            if _looks_like_refusal_response(out):
                return IMAGE_EXPLAINER_FAILURE_LINE
            return out
        raw_fb = str(parsed.get("raw", "")).strip()
        if _looks_like_refusal_response(raw_fb):
            return IMAGE_EXPLAINER_FAILURE_LINE
        return raw_fb
    summary = parsed.get("summary")
    if summary:
        out = str(summary).strip()
        if _looks_like_refusal_response(out):
            return IMAGE_EXPLAINER_FAILURE_LINE
        return out
    raw_fallback = str(parsed.get("raw", "")).strip()
    if _looks_like_refusal_response(raw_fallback):
        return IMAGE_EXPLAINER_FAILURE_LINE
    return raw_fallback




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


def normalize_multi_prompts(value: Any) -> List[str]:
    """Normalize multi_prompts from None, str, or list into a deduped ordered list."""
    if value is None:
        return []
    if isinstance(value, str):
        return _normalize_multi_prompts_list([value])
    if isinstance(value, list):
        return _normalize_multi_prompts_list(value)
    return _normalize_multi_prompts_list([str(value)])


def merge_multi_prompts(sources: Any) -> List[str]:
    """Merge multi_prompts from dict/list/string sources, preserving order and deduping."""
    merged: List[str] = []
    if not sources:
        return []
    if not isinstance(sources, list):
        sources = [sources]
    for source in sources:
        if source is None:
            continue
        if isinstance(source, str):
            merged.extend(normalize_multi_prompts(source))
            continue
        if isinstance(source, list):
            merged.extend(normalize_multi_prompts(source))
            continue
        if isinstance(source, dict):
            if source.get("multi_prompts") is not None:
                merged.extend(normalize_multi_prompts(source.get("multi_prompts")))
            meta = source.get("meta")
            if isinstance(meta, dict) and meta.get("multi_prompts") is not None:
                merged.extend(normalize_multi_prompts(meta.get("multi_prompts")))
    return normalize_multi_prompts(merged)


def is_enable_multi_prompts(kwargs: Optional[Dict[str, Any]] = None, default: bool = True) -> bool:
    """Return whether extract/segment/chunk should generate multi_prompts."""
    if not isinstance(kwargs, dict):
        return default
    value = kwargs.get("enable_multi_prompts", default)
    if value is None:
        return default
    return bool(value)


def attach_multi_prompts_meta(record: Dict[str, Any], prompts: Any) -> None:
    """Attach normalized multi_prompts into record meta without clobbering other keys."""
    meta = record.setdefault("meta", {})
    meta["multi_prompts"] = normalize_multi_prompts(prompts)


def attach_segment_multi_prompts_meta(segment_record: Dict[str, Any], source_paras: List[Dict[str, Any]]) -> None:
    """Merge upstream unit meta.multi_prompts into a segment record."""
    attach_multi_prompts_meta(segment_record, merge_multi_prompts(source_paras))


def attach_segment_multi_prompts_by_orders(
        segments: List[Dict[str, Any]],
        unit_paras: List[Dict[str, Any]],
        orders_key: str = 'orders',
    ) -> None:
    """Merge unit meta.multi_prompts into segments using tracked unit orders."""
    if not segments or not unit_paras:
        return
    para_by_order: Dict[Any, Dict[str, Any]] = {}
    for para in unit_paras:
        order = para.get('order')
        if order is not None:
            para_by_order[order] = para
    for seg in segments:
        orders = seg.get(orders_key)
        if isinstance(orders, list) and orders:
            sources = [para_by_order[o] for o in orders if o in para_by_order]
        elif len(segments) == 1:
            sources = unit_paras
        else:
            sources = []
        if sources:
            attach_segment_multi_prompts_meta(seg, sources)




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




def _apply_document_brief_prefix(template: str, metadata: Optional[Dict[str, Any]]) -> str:
    """將 metadata['document_brief'] 置於圖像說明 prompt 前綴（若有）。"""
    if not template:
        return ""
    if not isinstance(metadata, dict):
        return template
    brief = metadata.get("document_brief")
    if not brief or not str(brief).strip():
        return template
    return f"文件主旨／背景：{str(brief).strip()}\n\n{template}"


def _load_image_explainer_prompt_template(
        config: Optional[Dict[str, Any]],
        log: logging.Logger,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> str:
    """載入 image_explainer prompt 模板，失敗時回退內建預設；若有 document_brief 則加前綴。"""
    default_template = (
        "你是圖像內容分析器。請嚴格只輸出一個 JSON 物件，不要輸出 markdown code block、不要輸出額外說明。"
        "JSON 欄位固定為：{\"has_table\": <true|false>, \"summary\": \"...\", \"markdown_table\": \"...\"}。"
        "規則："
        "1) 先對齊上文下文的文件主題或章節主旨，說明此圖如何支撐該主題（用途、結論或讀圖重點），避免只描述底色、漸層或框線等與主題無關的視覺外觀。"
        "2) 若圖片中文字量明顯偏多（如合約條款、段落、清單、表單說明），且不屬於表格型圖片，請優先做逐字轉寫並完整保留原文內容；summary 欄位需放轉寫結果（可含換行）。僅在文字量少、難以辨識或非文字型圖片時，summary 才改為重點摘要。"
        "3) 若圖片中有條列式文案（例如 1. / 1.1 / 2.、項次、符號清單），請逐條列出且保持原有層次與順序，不可合併、不可省略。"
        "4) 若圖為報表／儀表板／統計或座標 xy chart，請在 summary 補充可辨識的標題、圖例、座標軸意義、時間區間、排名、關鍵數值與趨勢；看不清處可標註不確定。"
        "5) 先判斷圖片是否含表格，有則 has_table=true，否則 false。"
        "6) 若 has_table=true，markdown_table 必須輸出圖片中前 10 行前 10 欄的 markdown 表格（含表頭與分隔線）；若不足 10 行/欄則輸出實際可見範圍。此情況 summary 請限制為一句（<=80字），禁止貼出大量儲存格內容或公式。"
        "7) 若 has_table=false，markdown_table 請輸出空字串。"
        "8) 如果是空白表格範例，而且有明顯欄列標題，請在 summary 說明欄位用途，markdown_table 仍輸出可見表頭/欄位。但如果有內容的儲存格分布四散，請說明這份空白表格可能要如何填寫。"
    )
    cfg = config or {}
    base = default_template
    try:
        project_root = Path(__file__).resolve().parent.parent.parent
        prompt_path = cfg.get('prompt_path', None) if isinstance(cfg, dict) else None
        prompt_file = Path(prompt_path) / 'image_explainer.json' if prompt_path else (project_root / 'prompt' / 'image_explainer.json')

        if not prompt_file.exists():
            log.warning(f"[_analyze_images_via_batch] prompt 檔案不存在，使用內建預設: {prompt_file}")
        else:
            with open(prompt_file, 'r', encoding='utf-8-sig') as f:
                raw = f.read().strip()
            if not raw:
                log.warning(f"[_analyze_images_via_batch] prompt 檔案為空，使用內建預設: {prompt_file}")
            else:
                loaded = json.loads(raw)
                if isinstance(loaded, str):
                    template = loaded.strip()
                    base = template or default_template
                elif isinstance(loaded, dict):
                    picked = ""
                    for key in ('image_prompt_template', 'prompt_template', 'user_prompt_template'):
                        val = loaded.get(key)
                        if val is not None:
                            picked = str(val).strip()
                            if picked:
                                break
                    base = picked if picked else default_template
                else:
                    base = default_template
    except Exception as e:
        log.warning(f"[_analyze_images_via_batch] 載入 image_explainer prompt 失敗，使用內建預設: {e}")
        base = default_template

    return _apply_document_brief_prefix(base, metadata)


def _apply_image_analysis_to_segments(
        image_analysis_map: Dict[str, Any],
        segments: List[Dict[str, Any]],
        text_keys: List[str],
        placeholder_pattern: str,
        placeholder_replacements: List[str],
        logger: Optional[logging.Logger] = None
    ) -> None:
    if not image_analysis_map:
        return
    placeholder_re = re.compile(placeholder_pattern)
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
            summary = summary.strip()
            if summary == IMAGE_EXPLAINER_FAILURE_LINE:
                annotation_line = IMAGE_EXPLAINER_FAILURE_LINE
            else:
                hash_suffix = ""
                image_hash = analysis.get("image_hash")
                if isinstance(image_hash, str) and image_hash.strip():
                    hash_suffix = f"|hash={image_hash.strip()}"
                if analysis.get("annotation_type") == "supplemental_object":
                    annotation_line = f"[補充物件內容|id={iid}{hash_suffix}] {summary}\n"
                else:
                    annotation_line = f"[圖像說明|id={iid}{hash_suffix}] {summary}\n"
            if annotation_line in current_text:
                continue
            placeholders = _build_placeholders(iid_key, placeholder_replacements)
            for ph in placeholders:
                if ph in current_text:
                    current_text = current_text.replace(ph, annotation_line)
            if iid_key not in appended_ids:
                segment.setdefault("image_analysis", []).append(analysis)
                appended_ids.add(iid_key)
        segment[text_key] = current_text


def _get_text(segment: Dict[str, Any], text_keys: List[str]) -> str:
    for key in text_keys:
        if key in segment and segment.get(key):
            value = segment.get(key)
            return value if isinstance(value, str) else str(value)
    return ""

def _build_placeholders(image_id_str: str, placeholder_replacements: List[str]) -> List[str]:
    placeholders = []
    for template in placeholder_replacements:
        if "{image_id}" in template:
            placeholders.append(template.format(image_id=image_id_str))
        else:
            placeholders.append(template)
    return placeholders


def _compute_image_hash_from_metadata_row(image_row: Any) -> str:
    if not isinstance(image_row, dict):
        return ""
    raw_bytes: Optional[bytes] = None
    base64_value = image_row.get("base64")
    if isinstance(base64_value, str) and base64_value.strip():
        try:
            raw_bytes = base64.b64decode(base64_value.strip(), validate=False)
        except Exception:
            raw_bytes = None
    if raw_bytes is None:
        image_url = image_row.get("image_url")
        if isinstance(image_url, str) and image_url.startswith("data:"):
            match = re.match(r'^data:[^;,]+;base64,(.*)$', image_url, flags=re.IGNORECASE | re.DOTALL)
            if match:
                try:
                    raw_bytes = base64.b64decode((match.group(1) or "").strip(), validate=False)
                except Exception:
                    raw_bytes = None
    if not raw_bytes:
        return ""
    return hashlib.sha256(raw_bytes).hexdigest()[:16]


def _page_image_analysis_context_for_placeholder(
        metadata: Optional[Dict[str, Any]],
        image_id: int,
        images: List[Any],
    ) -> str:
    """依圖所屬頁碼，帶入 metadata['page_image_analysis'] 的整頁預覽摘要。"""
    if not isinstance(metadata, dict):
        return ""
    pia = metadata.get("page_image_analysis")
    if not isinstance(pia, dict) or not pia:
        return ""
    if not isinstance(images, list) or image_id < 0 or image_id >= len(images):
        return ""
    img_row = images[image_id]
    if not isinstance(img_row, dict):
        return ""
    pg = img_row.get("page")
    if pg is None:
        return ""
    try:
        pg_key = str(int(pg))
    except Exception:
        return ""
    entry = pia.get(pg_key)
    if not isinstance(entry, dict):
        return ""
    summ = entry.get("summary")
    if summ is None or not str(summ).strip():
        return ""
    return f"本頁整頁視覺脈絡（整頁預覽摘要）：{str(summ).strip()}"




def _load_image_explainer_prompt_template(
        config: Optional[Dict[str, Any]],
        log: logging.Logger,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> str:
    """載入 image_explainer prompt 模板，失敗時回退內建預設；若有 document_brief 則加前綴。"""
    default_template = (
        "你是圖像內容分析器。請嚴格只輸出一個 JSON 物件，不要輸出 markdown code block、不要輸出額外說明。"
        "JSON 欄位固定為：{\"has_table\": <true|false>, \"summary\": \"...\", \"markdown_table\": \"...\"}。"
        "規則："
        "1) 先對齊上文下文的文件主題或章節主旨，說明此圖如何支撐該主題（用途、結論或讀圖重點），避免只描述底色、漸層或框線等與主題無關的視覺外觀。"
        "2) 若圖片中文字量明顯偏多（如合約條款、段落、清單、表單說明），且不屬於表格型圖片，請優先做逐字轉寫並完整保留原文內容；summary 欄位需放轉寫結果（可含換行）。僅在文字量少、難以辨識或非文字型圖片時，summary 才改為重點摘要。"
        "3) 若圖片中有條列式文案（例如 1. / 1.1 / 2.、項次、符號清單），請逐條列出且保持原有層次與順序，不可合併、不可省略。"
        "4) 若圖為報表／儀表板／統計或座標 xy chart，請在 summary 補充可辨識的標題、圖例、座標軸意義、時間區間、排名、關鍵數值與趨勢；看不清處可標註不確定。"
        "5) 先判斷圖片是否含表格，有則 has_table=true，否則 false。"
        "6) 若 has_table=true，markdown_table 必須輸出圖片中前 10 行前 10 欄的 markdown 表格（含表頭與分隔線）；若不足 10 行/欄則輸出實際可見範圍。此情況 summary 請限制為一句（<=80字），禁止貼出大量儲存格內容或公式。"
        "7) 若 has_table=false，markdown_table 請輸出空字串。"
        "8) 如果是空白表格範例，而且有明顯欄列標題，請在 summary 說明欄位用途，markdown_table 仍輸出可見表頭/欄位。但如果有內容的儲存格分布四散，請說明這份空白表格可能要如何填寫。"
    )
    cfg = config or {}
    base = default_template
    try:
        project_root = Path(PROJECT_ROOT)
        prompt_path = cfg.get('prompt_path', None) if isinstance(cfg, dict) else None
        prompt_file = Path(prompt_path) / 'image_explainer.json' if prompt_path else (project_root / 'prompt' / 'image_explainer.json')

        if not prompt_file.exists():
            log.warning(f"[_analyze_images_via_batch] prompt 檔案不存在，使用內建預設: {prompt_file}")
        else:
            with open(prompt_file, 'r', encoding='utf-8-sig') as f:
                raw = f.read().strip()
            if not raw:
                log.warning(f"[_analyze_images_via_batch] prompt 檔案為空，使用內建預設: {prompt_file}")
            else:
                loaded = json.loads(raw)
                if isinstance(loaded, str):
                    template = loaded.strip()
                    base = template or default_template
                elif isinstance(loaded, dict):
                    picked = ""
                    for key in ('image_prompt_template', 'prompt_template', 'user_prompt_template'):
                        val = loaded.get(key)
                        if val is not None:
                            picked = str(val).strip()
                            if picked:
                                break
                    base = picked if picked else default_template
                else:
                    base = default_template
    except Exception as e:
        log.warning(f"[_analyze_images_via_batch] 載入 image_explainer prompt 失敗，使用內建預設: {e}")
        base = default_template

    return _apply_document_brief_prefix(base, metadata)


def _apply_image_analysis_to_segments(
        image_analysis_map: Dict[str, Any],
        segments: List[Dict[str, Any]],
        text_keys: List[str],
        placeholder_pattern: str,
        placeholder_replacements: List[str],
        logger: Optional[logging.Logger] = None
    ) -> None:
    if not image_analysis_map:
        return
    placeholder_re = re.compile(placeholder_pattern)
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
            summary = summary.strip()
            if summary == IMAGE_EXPLAINER_FAILURE_LINE:
                annotation_line = IMAGE_EXPLAINER_FAILURE_LINE
            else:
                hash_suffix = ""
                image_hash = analysis.get("image_hash")
                if isinstance(image_hash, str) and image_hash.strip():
                    hash_suffix = f"|hash={image_hash.strip()}"
                if analysis.get("annotation_type") == "supplemental_object":
                    annotation_line = f"[補充物件內容|id={iid}{hash_suffix}] {summary}\n"
                else:
                    annotation_line = f"[圖像說明|id={iid}{hash_suffix}] {summary}\n"
            if annotation_line in current_text:
                continue
            placeholders = _build_placeholders(iid_key, placeholder_replacements)
            for ph in placeholders:
                if ph in current_text:
                    current_text = current_text.replace(ph, annotation_line)
            if iid_key not in appended_ids:
                segment.setdefault("image_analysis", []).append(analysis)
                appended_ids.add(iid_key)
        segment[text_key] = current_text


def _get_text(segment: Dict[str, Any], text_keys: List[str]) -> str:
    for key in text_keys:
        if key in segment and segment.get(key):
            value = segment.get(key)
            return value if isinstance(value, str) else str(value)
    return ""

def _build_placeholders(image_id_str: str, placeholder_replacements: List[str]) -> List[str]:
    placeholders = []
    for template in placeholder_replacements:
        if "{image_id}" in template:
            placeholders.append(template.format(image_id=image_id_str))
        else:
            placeholders.append(template)
    return placeholders


def _compute_image_hash_from_metadata_row(image_row: Any) -> str:
    if not isinstance(image_row, dict):
        return ""
    raw_bytes: Optional[bytes] = None
    base64_value = image_row.get("base64")
    if isinstance(base64_value, str) and base64_value.strip():
        try:
            raw_bytes = base64.b64decode(base64_value.strip(), validate=False)
        except Exception:
            raw_bytes = None
    if raw_bytes is None:
        image_url = image_row.get("image_url")
        if isinstance(image_url, str) and image_url.startswith("data:"):
            match = re.match(r'^data:[^;,]+;base64,(.*)$', image_url, flags=re.IGNORECASE | re.DOTALL)
            if match:
                try:
                    raw_bytes = base64.b64decode((match.group(1) or "").strip(), validate=False)
                except Exception:
                    raw_bytes = None
    if not raw_bytes:
        return ""
    return hashlib.sha256(raw_bytes).hexdigest()[:16]


def _page_image_analysis_context_for_placeholder(
        metadata: Optional[Dict[str, Any]],
        image_id: int,
        images: List[Any],
    ) -> str:
    """依圖所屬頁碼，帶入 metadata['page_image_analysis'] 的整頁預覽摘要。"""
    if not isinstance(metadata, dict):
        return ""
    pia = metadata.get("page_image_analysis")
    if not isinstance(pia, dict) or not pia:
        return ""
    if not isinstance(images, list) or image_id < 0 or image_id >= len(images):
        return ""
    img_row = images[image_id]
    if not isinstance(img_row, dict):
        return ""
    pg = img_row.get("page")
    if pg is None:
        return ""
    try:
        pg_key = str(int(pg))
    except Exception:
        return ""
    entry = pia.get(pg_key)
    if not isinstance(entry, dict):
        return ""
    summ = entry.get("summary")
    if summ is None or not str(summ).strip():
        return ""
    return f"本頁整頁視覺脈絡（整頁預覽摘要）：{str(summ).strip()}"




def _apply_image_analysis_to_segments(
        image_analysis_map: Dict[str, Any],
        segments: List[Dict[str, Any]],
        text_keys: List[str],
        placeholder_pattern: str,
        placeholder_replacements: List[str],
        logger: Optional[logging.Logger] = None
    ) -> None:
    if not image_analysis_map:
        return
    placeholder_re = re.compile(placeholder_pattern)
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
            summary = summary.strip()
            if summary == IMAGE_EXPLAINER_FAILURE_LINE:
                annotation_line = IMAGE_EXPLAINER_FAILURE_LINE
            else:
                hash_suffix = ""
                image_hash = analysis.get("image_hash")
                if isinstance(image_hash, str) and image_hash.strip():
                    hash_suffix = f"|hash={image_hash.strip()}"
                if analysis.get("annotation_type") == "supplemental_object":
                    annotation_line = f"[補充物件內容|id={iid}{hash_suffix}] {summary}\n"
                else:
                    annotation_line = f"[圖像說明|id={iid}{hash_suffix}] {summary}\n"
            if annotation_line in current_text:
                continue
            placeholders = _build_placeholders(iid_key, placeholder_replacements)
            for ph in placeholders:
                if ph in current_text:
                    current_text = current_text.replace(ph, annotation_line)
            if iid_key not in appended_ids:
                segment.setdefault("image_analysis", []).append(analysis)
                appended_ids.add(iid_key)
        segment[text_key] = current_text


def _get_text(segment: Dict[str, Any], text_keys: List[str]) -> str:
    for key in text_keys:
        if key in segment and segment.get(key):
            value = segment.get(key)
            return value if isinstance(value, str) else str(value)
    return ""

def _build_placeholders(image_id_str: str, placeholder_replacements: List[str]) -> List[str]:
    placeholders = []
    for template in placeholder_replacements:
        if "{image_id}" in template:
            placeholders.append(template.format(image_id=image_id_str))
        else:
            placeholders.append(template)
    return placeholders


def _compute_image_hash_from_metadata_row(image_row: Any) -> str:
    if not isinstance(image_row, dict):
        return ""
    raw_bytes: Optional[bytes] = None
    base64_value = image_row.get("base64")
    if isinstance(base64_value, str) and base64_value.strip():
        try:
            raw_bytes = base64.b64decode(base64_value.strip(), validate=False)
        except Exception:
            raw_bytes = None
    if raw_bytes is None:
        image_url = image_row.get("image_url")
        if isinstance(image_url, str) and image_url.startswith("data:"):
            match = re.match(r'^data:[^;,]+;base64,(.*)$', image_url, flags=re.IGNORECASE | re.DOTALL)
            if match:
                try:
                    raw_bytes = base64.b64decode((match.group(1) or "").strip(), validate=False)
                except Exception:
                    raw_bytes = None
    if not raw_bytes:
        return ""
    return hashlib.sha256(raw_bytes).hexdigest()[:16]


def _page_image_analysis_context_for_placeholder(
        metadata: Optional[Dict[str, Any]],
        image_id: int,
        images: List[Any],
    ) -> str:
    """依圖所屬頁碼，帶入 metadata['page_image_analysis'] 的整頁預覽摘要。"""
    if not isinstance(metadata, dict):
        return ""
    pia = metadata.get("page_image_analysis")
    if not isinstance(pia, dict) or not pia:
        return ""
    if not isinstance(images, list) or image_id < 0 or image_id >= len(images):
        return ""
    img_row = images[image_id]
    if not isinstance(img_row, dict):
        return ""
    pg = img_row.get("page")
    if pg is None:
        return ""
    try:
        pg_key = str(int(pg))
    except Exception:
        return ""
    entry = pia.get(pg_key)
    if not isinstance(entry, dict):
        return ""
    summ = entry.get("summary")
    if summ is None or not str(summ).strip():
        return ""
    return f"本頁整頁視覺脈絡（整頁預覽摘要）：{str(summ).strip()}"




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

    cfg = load_config(config)
    log = logger or logging.getLogger(__name__)

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




def _get_text(segment: Dict[str, Any], text_keys: List[str]) -> str:
    for key in text_keys:
        if key in segment and segment.get(key):
            value = segment.get(key)
            return value if isinstance(value, str) else str(value)
    return ""

def _build_placeholders(image_id_str: str, placeholder_replacements: List[str]) -> List[str]:
    placeholders = []
    for template in placeholder_replacements:
        if "{image_id}" in template:
            placeholders.append(template.format(image_id=image_id_str))
        else:
            placeholders.append(template)
    return placeholders


def _compute_image_hash_from_metadata_row(image_row: Any) -> str:
    if not isinstance(image_row, dict):
        return ""
    raw_bytes: Optional[bytes] = None
    base64_value = image_row.get("base64")
    if isinstance(base64_value, str) and base64_value.strip():
        try:
            raw_bytes = base64.b64decode(base64_value.strip(), validate=False)
        except Exception:
            raw_bytes = None
    if raw_bytes is None:
        image_url = image_row.get("image_url")
        if isinstance(image_url, str) and image_url.startswith("data:"):
            match = re.match(r'^data:[^;,]+;base64,(.*)$', image_url, flags=re.IGNORECASE | re.DOTALL)
            if match:
                try:
                    raw_bytes = base64.b64decode((match.group(1) or "").strip(), validate=False)
                except Exception:
                    raw_bytes = None
    if not raw_bytes:
        return ""
    return hashlib.sha256(raw_bytes).hexdigest()[:16]


def _page_image_analysis_context_for_placeholder(
        metadata: Optional[Dict[str, Any]],
        image_id: int,
        images: List[Any],
    ) -> str:
    """依圖所屬頁碼，帶入 metadata['page_image_analysis'] 的整頁預覽摘要。"""
    if not isinstance(metadata, dict):
        return ""
    pia = metadata.get("page_image_analysis")
    if not isinstance(pia, dict) or not pia:
        return ""
    if not isinstance(images, list) or image_id < 0 or image_id >= len(images):
        return ""
    img_row = images[image_id]
    if not isinstance(img_row, dict):
        return ""
    pg = img_row.get("page")
    if pg is None:
        return ""
    try:
        pg_key = str(int(pg))
    except Exception:
        return ""
    entry = pia.get(pg_key)
    if not isinstance(entry, dict):
        return ""
    summ = entry.get("summary")
    if summ is None or not str(summ).strip():
        return ""
    return f"本頁整頁視覺脈絡（整頁預覽摘要）：{str(summ).strip()}"




def analyze_images_via_batch_common(
        segments: List[Dict[str, Any]],
        metadata: Optional[Dict[str, Any]],
        *,
        llm_provider: str = 'openai',
        llm_model: str = 'gpt4o_chat',
        llm_base_url: Optional[str] = None,
        enable_image_llm: bool = True,
        image_context_window: int = 200, # 可能影響到速度
        max_images_per_batch: int = 50,
        image_prompt_template: Optional[str] = None,
        placeholder_pattern: Optional[str] = None,
        placeholder_replacements: Optional[Union[str, List[str]]] = None,
        text_keys: Optional[List[str]] = None,
        config: Optional[Dict[str, Any]] = None,
        logger: Optional[logging.Logger] = None,
        parser_type: Optional[str] = None,
        fusion_overrides: Optional[Dict[str, Any]] = None,
        analysis_only: bool = False,
    ) -> Dict[str, Any]:
    """
    Shared image analysis for parsers. Scans placeholders, calls /chat/batch,
    and inserts summaries back into segments and metadata.
    """
    timing_t0 = time.perf_counter()
    timing_info: Dict[str, float] = {
        "prepare_tasks_seconds": 0.0,
        "batch_api_seconds": 0.0,
        "apply_segments_seconds": 0.0,
        "page_fusion_seconds": 0.0,
        "total_seconds": 0.0,
    }
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
        image_prompt_template = _load_image_explainer_prompt_template(cfg, log, metadata)
    else:
        image_prompt_template = _apply_document_brief_prefix(image_prompt_template, metadata)

    if not placeholder_pattern:
        placeholder_pattern = r'\[IMAGE_PLACEHOLDER_(\d+)\]'
    placeholder_re = re.compile(placeholder_pattern)

    if not text_keys:
        text_keys = ['unit_text', 'call_prompt', 'text']

    if placeholder_replacements is None:
        placeholder_replacements = ["[IMAGE_PLACEHOLDER_{image_id}]"]
    elif isinstance(placeholder_replacements, str):
        placeholder_replacements = [placeholder_replacements]
        
    image_tasks: List[Dict[str, Any]] = []
    seen_image_ids: Set[int] = set()
    placeholder_ref_total = 0

    t_stage = time.perf_counter()
    for seg_idx, segment in enumerate(segments):
        text = _get_text(segment, text_keys)
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
                prev_text = _get_text(segments[seg_idx - 1], text_keys)
                front_text = prev_text[-image_context_window:] if len(prev_text) > image_context_window else prev_text
            if seg_idx < len(segments) - 1:
                next_text = _get_text(segments[seg_idx + 1], text_keys)
                back_text = next_text[:image_context_window] if len(next_text) > image_context_window else next_text

            if front_text or back_text:
                prompt = f"{image_prompt_template}\n上文：{front_text}\n下文：{back_text}"
            else:
                prompt = image_prompt_template
            page_ctx = _page_image_analysis_context_for_placeholder(metadata, image_id, images)
            if page_ctx:
                prompt = f"{prompt}\n{page_ctx}"

            img_data = images[image_id]
            image_url = None
            if isinstance(img_data, dict):
                shape_ctx = str(img_data.get("context_text") or "").strip()
                if shape_ctx:
                    prompt = f"{prompt}\n同頁可讀文字參考：{shape_ctx}"
                if str(img_data.get("source") or "") == "slide_screenshot_fallback":
                    prompt = (
                        f"{prompt}\n請僅補充『同頁可讀文字參考』未涵蓋的物件資訊；"
                        f"不要重述已出現內容，若無新增資訊請回覆空字串。"
                    )
                if 'base64' in img_data:
                    mime = img_data.get('mime', 'image/jpeg')
                    base64_str = img_data['base64']
                    image_url = f"data:{mime};base64,{base64_str}"
                elif 'image_url' in img_data:
                    image_url = img_data['image_url']

            if image_url:
                image_url, converted = _normalize_image_data_url_for_llm(image_url, logger=log)
                if converted:
                    log.debug(f"[_analyze_images_via_batch] 圖像 ID {image_id} 已轉為 PNG 後送出")
                if image_url.startswith("data:"):
                    m = re.match(r'^data:([^;,]+);base64,', image_url, flags=re.IGNORECASE)
                    mime_after = (m.group(1).strip().lower() if m else "")
                    if mime_after and mime_after not in {"image/jpeg", "image/png", "image/gif", "image/webp"}:
                        log.warning(f"[_analyze_images_via_batch] 圖像 ID {image_id} 格式仍不支援，略過: {mime_after}")
                        image_url = None

            if not image_url:
                log.warning(f"[_analyze_images_via_batch] 圖像 ID {image_id} 無法取得圖像資料")
                continue

            image_tasks.append({
                'image_id': image_id,
                'image_url': image_url,
                'prompt': prompt,
                'segment_idx': seg_idx,
                'sheet_name': segment.get('sheet_name'),
                'placeholder_match': match,
                'image_hash': _compute_image_hash_from_metadata_row(img_data),
                'submitted_at': None,
            })
            if isinstance(img_data, dict) and image_tasks[-1].get('image_hash'):
                img_data['image_hash'] = image_tasks[-1]['image_hash']
    timing_info["prepare_tasks_seconds"] = round(time.perf_counter() - t_stage, 3)

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
            t_batch = time.perf_counter()
            for task in batch_tasks:
                task['submitted_at'] = t_batch
            payload = {
                'prompts': [task['prompt'] for task in batch_tasks],
                'provider': llm_provider,
                'model': llm_model,
                # 'max_tokens': 12000,
                'max_tokens': 120000,
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
                            elapsed = round(time.perf_counter() - (task.get('submitted_at') or t_batch), 3)
                            log.warning(f"[_analyze_images_via_batch] 圖像 ID {task['image_id']} 分析失敗: {result_item['error']}")
                            image_analysis[str(task['image_id'])] = {
                                'error': result_item['error'],
                                'image_id': task['image_id'],
                                'segment_idx': task['segment_idx'],
                                'sheet_name': task.get('sheet_name'),
                                'image_hash': task.get('image_hash') or "",
                                'elapsed_seconds': elapsed,
                            }
                            log.info(f"[_analyze_images_via_batch] 圖像 ID {task['image_id']} 解析耗時={elapsed}s (error)")
                        else:
                            elapsed = round(time.perf_counter() - (task.get('submitted_at') or t_batch), 3)
                            result = result_item.get('result') or {}
                            output = result.get('output', '')
                            # 圖文解釋：後續在此銜接 has_table / md 表格檢查與重試（見 src.image_explainer_output）
                            output = finalize_image_explainer_text(output)
                            cost = result.get('current_cost_usd', 0)
                            usage = result.get('usage') or {}
                            total_tokens = usage.get('total_tokens', 0) if isinstance(usage, dict) else 0

                            prompt_tokens = usage.get('prompt_tokens', 0) if isinstance(usage, dict) else 0
                            completion_tokens = usage.get('completion_tokens', 0) if isinstance(usage, dict) else 0
                            provider_key = str(llm_provider or "openai").strip().lower()
                            provider_name = "OpenAI" if provider_key == "openai" else str(llm_provider or "Unknown")

                            image_analysis[str(task['image_id'])] = {
                                'summary': output,
                                'prompt_used': task['prompt'],
                                'model': llm_model,
                                'image_hash': task.get('image_hash') or "",
                                'cost': cost,
                                'tokens': total_tokens,
                                'prompt_tokens': prompt_tokens,
                                'completion_tokens': completion_tokens,
                                'usage': {
                                    'cost_types': {
                                        provider_name: {
                                            'tokens_in': prompt_tokens,
                                            'tokens_out': completion_tokens,
                                            'model': llm_model
                                        }
                                    }
                                },
                                'image_id': task['image_id'],
                                'segment_idx': task['segment_idx'],
                                'sheet_name': task.get('sheet_name'),
                                'elapsed_seconds': elapsed,
                            }
                            log.info(f"[_analyze_images_via_batch] 圖像 ID {task['image_id']} 解析耗時={elapsed}s")
                    else:
                        elapsed = round(time.perf_counter() - (task.get('submitted_at') or t_batch), 3)
                        log.warning(f"[_analyze_images_via_batch] 圖像 ID {task['image_id']} 沒有對應的回覆")
                        image_analysis[str(task['image_id'])] = {
                            'error': 'No response',
                            'image_id': task['image_id'],
                            'segment_idx': task['segment_idx'],
                            'sheet_name': task.get('sheet_name'),
                            'image_hash': task.get('image_hash') or "",
                            'elapsed_seconds': elapsed,
                        }
                        log.info(f"[_analyze_images_via_batch] 圖像 ID {task['image_id']} 解析耗時={elapsed}s (no_response)")

                batch_cost = batch_data.get('batch_cost_usd', 0)
                batch_token_summary = batch_data.get('batch_token_summary') or {}
                if not isinstance(batch_token_summary, dict):
                    batch_token_summary = {}
                log.info(f"[_analyze_images_via_batch] 批次 {batch_start//max_images_per_batch + 1} 完成: 費用={batch_cost}, tokens={batch_token_summary.get('total_tokens', 0)}")

                timing_info["batch_api_seconds"] += (time.perf_counter() - t_batch)
            except requests.RequestException as e:
                log.error(f"[_analyze_images_via_batch] 批次請求失敗: {e}")
                for task in batch_tasks:
                    elapsed = round(time.perf_counter() - (task.get('submitted_at') or t_batch), 3)
                    image_analysis[str(task['image_id'])] = {
                        'error': str(e),
                        'image_id': task['image_id'],
                        'segment_idx': task['segment_idx'],
                        'sheet_name': task.get('sheet_name'),
                        'image_hash': task.get('image_hash') or "",
                        'elapsed_seconds': elapsed,
                    }
                    log.info(f"[_analyze_images_via_batch] 圖像 ID {task['image_id']} 解析耗時={elapsed}s (request_exception)")
                timing_info["batch_api_seconds"] += (time.perf_counter() - t_batch)

        if image_analysis and isinstance(metadata, dict):
            metadata['image_analysis'] = image_analysis

            # Aggregate image-analysis usage for downstream consumers.
            provider_usage_map: Dict[str, Dict[str, Any]] = {}
            image_count = 0
            for ia in image_analysis.values():
                if not isinstance(ia, dict):
                    continue
                if ia.get("error"):
                    continue
                image_count += 1
                usage_obj = ia.get("usage") if isinstance(ia.get("usage"), dict) else {}
                cost_types = usage_obj.get("cost_types") if isinstance(usage_obj.get("cost_types"), dict) else {}
                for provider_name, provider_info in cost_types.items():
                    if not isinstance(provider_info, dict):
                        continue
                    tokens_in = int(provider_info.get("tokens_in", 0) or 0)
                    tokens_out = int(provider_info.get("tokens_out", 0) or 0)
                    model_name = provider_info.get("model")
                    entry = provider_usage_map.setdefault(
                        str(provider_name),
                        {"tokens_in": 0, "tokens_out": 0, "model": model_name}
                    )
                    entry["tokens_in"] += tokens_in
                    entry["tokens_out"] += tokens_out
                    if not entry.get("model") and model_name:
                        entry["model"] = model_name

            metadata["usage"] = {
                "cost_types": provider_usage_map,
                "image_count": image_count
            }
            try:
                from src.tb11_usage_tracker import accumulate_usage_payload
                accumulate_usage_payload(metadata["usage"], request_count=image_count)
            except Exception:
                pass

    t_stage = time.perf_counter()
    for iid, analysis in (image_analysis or {}).items():
        try:
            ii = int(iid)
        except Exception:
            continue
        if 0 <= ii < len(images):
            img_row = images[ii]
            if isinstance(img_row, dict) and str(img_row.get("source") or "") == "slide_screenshot_fallback":
                if isinstance(analysis, dict):
                    analysis["annotation_type"] = "supplemental_object"

    if analysis_only:
        timing_info["apply_segments_seconds"] = 0.0
        timing_info["page_fusion_seconds"] = 0.0
        timing_info["total_seconds"] = round(time.perf_counter() - timing_t0, 3)
        if isinstance(metadata, dict):
            metadata["image_analysis_timing"] = timing_info
        return {
            'image_analysis': image_analysis,
            'total_analyzed': len(image_analysis),
            'total_errors': sum(1 for v in image_analysis.values() if isinstance(v, dict) and 'error' in v),
            'timing': timing_info,
            'analysis_only': True,
        }

    _apply_image_analysis_to_segments(
        image_analysis_map=image_analysis,
        segments=segments,
        text_keys=text_keys,
        placeholder_pattern=placeholder_pattern,
        placeholder_replacements=placeholder_replacements,
        logger=log,
    )

    timing_info["page_fusion_seconds"] = 0.0

    log.info(f"[_analyze_images_via_batch] 完成圖像分析，共 {len(image_analysis)} 個結果")

    timing_info["apply_segments_seconds"] = round(time.perf_counter() - t_stage, 3)
    timing_info["batch_api_seconds"] = round(timing_info["batch_api_seconds"], 3)
    timing_info["total_seconds"] = round(time.perf_counter() - timing_t0, 3)
    if isinstance(metadata, dict):
        metadata["image_analysis_timing"] = timing_info
    return {
        'image_analysis': image_analysis,
        'total_analyzed': len(image_analysis),
        'total_errors': sum(1 for v in image_analysis.values() if 'error' in v),
        'timing': timing_info,
    }




def batch_extract_multi_prompts(
        texts: List[str],
        *,
        llm_base_url: Optional[str] = None,
        prompt_config: Optional[Dict[str, Any]] = None,
        prompt_file: Optional[Union[str, Path]] = None,
        llm_provider: str = "remote",
        llm_model: str = "remote8b",
        min_text_length: int = 1,
        batch_size: int = None,
        timeout: int = 300,
    ) -> List[List[str]]:
    """
    Batch extract multi_prompts via /chat/batch using extract_keywords prompt config.
    Returns one keyword list per input text; falls back to empty lists on LLM failure.
    """
    if not texts:
        return []

    kw_config = _load_extract_keywords_prompt_config(prompt_config=prompt_config, prompt_file=prompt_file)
    if not kw_config:
        return [[] for _ in texts]

    system_prompt = kw_config.get("system_prompt", "")
    user_prompt_template = kw_config.get("user_prompt_template", "")
    generation_config = kw_config.get("generation_config", {})
    max_keywords = int(generation_config.get("max_keywords", 5))
    min_keyword_length = int(generation_config.get("min_keyword_length", 1))
    max_tokens = int(generation_config.get("max_new_tokens", 150))
    temperature = float(generation_config.get("temperature", 0.3))

    if not llm_base_url:
        llm_base_url = load_config().get("llm", {}).get("base_url", "http://10.1.3.127:7017")

    prompts: List[str] = []
    index_map: List[int] = []
    results: List[List[str]] = [[] for _ in texts]

    for idx, text in enumerate(texts):
        text_clean = (text or "").strip()
        if len(text_clean) < min_text_length:
            continue
        text_sample = text_clean[:5000]
        text_sample = text_sample.replace("，", ",").replace("：", ":").replace("；", ";")
        user_prompt = user_prompt_template.format(content=text_sample)
        prompts.append(user_prompt)
        index_map.append(idx)

    if not prompts:
        return results

    if batch_size is None:
        batch_size = _BATCH_SIZE_LIMIT

    batch_chat_url = f"{llm_base_url.rstrip('/')}/chat/batch"
    outputs: List[str] = []

    try:
        for start in range(0, len(prompts), max(1, batch_size)):
            batch_prompts = prompts[start:start + max(1, batch_size)]
            payload = {
                "prompts": batch_prompts,
                "provider": llm_provider,
                "model": llm_model,
                "max_tokens": max_tokens,
                "temperature": temperature,
                "system_prompt": system_prompt if system_prompt else None,
                "parallel": True,
                "max_batch_size": max(1, batch_size),
            }
            response = requests.post(batch_chat_url, json=payload, timeout=timeout)
            response.raise_for_status()
            batch_result = response.json()
            batch_items = batch_result.get("results", [])
            for item in batch_items:
                if item.get("error"):
                    outputs.append("")
                else:
                    result_obj = item.get("result", {})
                    outputs.append((result_obj.get("output", "") or "").strip())
    except Exception as e:
        logging.getLogger(__name__).warning(f"batch_extract_multi_prompts 失敗: {e}")
        return results

    for text_idx, output in zip(index_map, outputs):
        if not output:
            continue
        keywords = parse_keywords_from_text(
            output,
            max_keywords=max_keywords,
            min_keyword_length=min_keyword_length,
        )
        results[text_idx] = normalize_multi_prompts(keywords)

    return results


