#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM數學規律偵測器 - 使用語言模型偵測數字之間的數學規律
"""
import re
import os
import json
import requests
import openai
from typing import List, Dict, Any, Optional, Tuple
from dotenv import load_dotenv
import argparse
from pathlib import Path

from utils import LOGger

m_logfile = Path(__file__).resolve().parent.parent / "outputs" / "llm_math_rule_detector.log"
m_print = LOGger.addloger(logfile=str(m_logfile))


# 載入環境變數
load_dotenv()

def is_dollar_expr(s: str) -> bool:
    """Quick check for allowed token characters in $(a,b) / $i expressions."""
    return bool(re.match(r'^[\$\d\+\-\*\/\(\)\.,\s]+$', str(s).strip()))

def _detect_with_local_llm(
    clsInstance,
    values_matrix: List[List[Any]],
    row_names: List[str],
    column_names: List[str],
) -> Optional[Dict[str, Any]]:
    """使用TextProcessor服務偵測"""
    m_print(" 使用TextProcessor服務偵測數學規律")

    try:
        # 準備提示
        prompt_template = clsInstance.prompts.get('math_rule')
        if not prompt_template:
            m_print(" 找不到數學規律提示模板")
            return None
        
        # 構建完整的用戶提示（包含系統提示）
        full_prompt = (
            f"{prompt_template['system_prompt']}\n\n"
            f"{prompt_template['user_prompt'].format(values_matrix=values_matrix, row_names=row_names, column_names=column_names)}"
        )
        
        # 呼叫TextProcessor服務
        payload = {
            'prompt': full_prompt,
            'provider': clsInstance.textprocessor_provider,
            'model': clsInstance.textprocessor_model,
            'max_tokens': 2000,
            'temperature': 0.1
        }
        
        response = clsInstance.session.post(
            clsInstance.textprocessor_url,
            json=payload,
            timeout=120
        )
        
        if response.status_code != 200:
            m_print(f" TextProcessor API錯誤: {response.status_code} - {response.text}")
            return None
        
        try:
            # 確保回應編碼正確
            response.encoding = 'utf-8'
            result = response.json()
        except json.JSONDecodeError as e:
            m_print(f" TextProcessor回應JSON解析失敗: {e}")
            m_print(f" 原始回應: {response.text}")
            return None
        
        # 檢查TextProcessor回應格式
        # m_print(f" TextProcessor原始回應: {result}")
        
        # 檢查是否有錯誤
        if 'error' in result:
            m_print(f" TextProcessor回傳錯誤: {result['error']}")
            return None
        
        # 提取回應內容 - TextProcessor的主要內容在'output'欄位
        content = result.get('output')
        if not content:
            m_print(" TextProcessor回應為空")
            # 檢查是否有其他可能的內容欄位
            if 'response' in result:
                content = result.get('response')
                m_print(" 使用'response'欄位作為內容")
            elif 'message' in result:
                content = result.get('message')
                m_print(" 使用'message'欄位作為內容")
            else:
                m_print(" 找不到任何內容欄位")
                return None
        
        # 調試：檢查內容格式
        m_print(f" 提取的內容類型: {type(content)}")
        m_print(f" 提取的內容長度: {len(str(content))}")
        m_print(f" 內容前100字符: {str(content)[:100]}")
        
        # 確保內容是字串格式
        if not isinstance(content, str):
            content = str(content)
            m_print(" 將內容轉換為字串格式")
        
        # 記錄TextProcessor的元資訊
        provider = result.get('provider', 'unknown')
        model_alias = result.get('model_alias', 'unknown')
        post_id = result.get('post_id', 'unknown')
        timestamp = result.get('timestamp', 'unknown')
        
        m_print(f" TextProcessor Provider: {provider}")
        m_print(f" TextProcessor Model: {model_alias}")
        m_print(f" TextProcessor Post ID: {post_id}")
        
        # 記錄計價資訊
        usage_info = result.get('usage', {})
        billing_info = result.get('billing_summary', {})
        token_info = result.get('token_summary', {})
        
        # if usage_info:
        #     m_print(f" Token使用量: {usage_info}")
        # if billing_info:
        #     m_print(f" 計價資訊: {billing_info}")
        if token_info:
            m_print(f" Token摘要: {token_info}")
        
        # 解析回應
        parsed_result = clsInstance._parse_llm_response(content, 'textprocessor')
        
        # 將TextProcessor的完整資訊添加到結果中
        if parsed_result:
            parsed_result['textprocessor_usage'] = usage_info
            parsed_result['textprocessor_billing'] = billing_info
            parsed_result['textprocessor_token_summary'] = token_info
            parsed_result['textprocessor_post_id'] = post_id
            parsed_result['textprocessor_timestamp'] = timestamp
            parsed_result['textprocessor_provider'] = provider
            parsed_result['textprocessor_model_alias'] = model_alias
            parsed_result['textprocessor_raw_response'] = result
        
        return parsed_result
        
    except Exception as e:
        m_print(f" TextProcessor呼叫失敗: {e}")
        return None

class LLMMathRuleDetector:
    """LLM數學規律偵測器"""
    
    def __init__(self, prefer_local: bool = True, openai_model: str = "gpt35_chat"):
        """
        初始化偵測器
        
        Args:
            prefer_local: 是否優先使用TextProcessor服務
            openai_model: OpenAI模型名稱（預設: gpt35_chat
        """
        self.prefer_local = prefer_local
        self.openai_model = openai_model
        
        # TextProcessor服務設定
        self.textprocessor_url = os.getenv('TEXTPROCESSOR_URL', 'http://10.1.3.127:6017/chat')
        self.textprocessor_provider = os.getenv('TEXTPROCESSOR_PROVIDER', 'remote')
        self.textprocessor_model = os.getenv('TEXTPROCESSOR_MODEL', 'remote8b')
        
        # OpenAI設定（透過TextProcessor）
        self.openai_model = openai_model
        
        # 初始化客戶端
        self._init_clients()
        
        # 載入提示模板
        self.prompts = self._load_prompts()
    
    def _init_clients(self):
        """初始化LLM客戶端"""
        # TextProcessor服務session
        self.session = requests.Session()
        base_headers = {
            'Content-Type': 'application/json',
            'User-Agent': 'TextProcessor-Client/1.0'
        }
        self.session.headers.update(base_headers)
        m_print(" TextProcessor服務session已初始化")
    
    def _load_prompts(self) -> Dict[str, Any]:
        """載入提示模板"""
        prompts = {}
        prompt_dir = Path(__file__).parent.parent / 'prompts'
        
        try:
            prompt_file = prompt_dir / 'math_rule_detection.json'
            if prompt_file.exists():
                with open(prompt_file, 'r', encoding='utf-8-sig') as f:
                    prompts['math_rule'] = json.load(f)
                m_print(" 數學規律偵測提示模板已載入")
            else:
                m_print(f" 找不到提示模板檔案: {prompt_file}")
        except Exception as e:
            m_print(f" 載入提示模板失敗: {e}")
        
        return prompts

    def _normalize_matrix_inputs(
        self,
        values_matrix: Any,
        row_names: Optional[List[str]],
        column_names: Optional[List[str]],
    ) -> Tuple[List[List[Any]], List[str], List[str]]:
        """Normalize inputs into a 2D matrix with row/column names."""
        # Coerce to 2D list
        if not isinstance(values_matrix, list):
            matrix = []
        elif values_matrix and not isinstance(values_matrix[0], list):
            matrix = [[v] for v in values_matrix]
        else:
            matrix = values_matrix

        # Pad rows to uniform length
        max_cols = max((len(r) for r in matrix), default=0)
        matrix = [list(r) + [None] * (max_cols - len(r)) for r in matrix]

        # Normalize row names
        if not row_names or len(row_names) != len(matrix):
            row_names = [f"row_{i}" for i in range(len(matrix))]

        # Normalize column names
        if not column_names or len(column_names) != max_cols:
            column_names = [f"col_{i}" for i in range(max_cols)]

        return matrix, row_names, column_names

    def _count_numeric_cells(self, matrix: List[List[Any]]) -> int:
        count = 0
        for row in matrix:
            for v in row:
                if isinstance(v, (int, float)):
                    count += 1
        return count
    
    def detect_math_rules(
        self,
        values_matrix: List[List[Any]],
        row_names: List[str] = None,
        column_names: List[str] = None,
        use_openai: bool = False,
    ) -> Optional[Dict[str, Any]]:
        """
        偵測數值之間的數學規律
        
        Args:
            values: 數值列表
            column_names: 對應的欄位名稱
            use_openai: 是否強制使用OpenAI
            
        Returns:
            Dict: 偵測結果
        """
        values_matrix, row_names, column_names = self._normalize_matrix_inputs(
            values_matrix, row_names, column_names
        )
        if self._count_numeric_cells(values_matrix) < 2:
            m_print(" 需要至少2個數值才能偵測規律")
            return None

        m_print(f"偵測數學規律: matrix {len(values_matrix)}x{(len(values_matrix[0]) if values_matrix else 0)}")
        m_print(f"row_names: {row_names}")
        m_print(f"column_names: {column_names}")
        
        # 選擇LLM
        if use_openai or not self.prefer_local:
            result = self._detect_with_openai(values_matrix, row_names, column_names)
            if result:
                return result
            m_print(" TextProcessor OpenAI偵測失敗，嘗試TextProcessor 8B地端核心")
        
        # 使用TextProcessor 8B地端核心
        result = self._detect_with_textprocessor(values_matrix, row_names, column_names)
        if result:
            return result
        
        # 如果TextProcessor 8B地端核心失敗且還沒試過OpenAI
        if not use_openai and self.prefer_local:
            m_print(" TextProcessor 8B地端核心偵測失敗，嘗試TextProcessor OpenAI")
            result = self._detect_with_openai(values_matrix, row_names, column_names)
            if result:
                return result
        
        m_print(" 所有LLM偵測都失敗")
        return None
    
    def detect_math_rules_batch(
        self,
        batch_items: List[Dict[str, Any]],
        use_openai: bool = False,
        batch_size: int = 100
    ) -> List[Optional[Dict[str, Any]]]:
        """
        Batch detect math rules via TextProcessor /chat/batch.

        Args:
            batch_items: [{'values': [...], 'index_names': [...]}...]
            use_openai: route through TextProcessor OpenAI provider
            batch_size: chunk size for each /chat/batch request

        Returns:
            List[Optional[Dict]] aligned with input order.
        """
        if not batch_items:
            return []

        prompt_template = self.prompts.get('math_rule')
        if not prompt_template:
            m_print(" Prompt template missing for batch mode")
            return [None] * len(batch_items)

        results: List[Optional[Dict[str, Any]]] = [None] * len(batch_items)
        batch_url = self.textprocessor_url.rstrip('/')
        if batch_url.endswith('/chat'):
            batch_url = f"{batch_url}/batch"
        else:
            batch_url = f"{batch_url}/batch"

        provider = self.textprocessor_provider
        model = self.textprocessor_model
        llm_type = 'textprocessor'
        if use_openai or not self.prefer_local:
            provider = 'openai'
            model = self.openai_model
            if model == 'gpt-4o':
                model = 'gpt4o_chat'
            elif model == 'gpt-4':
                model = 'gpt4_chat'
            elif model == 'gpt-3.5-turbo':
                model = 'gpt35_chat'
            llm_type = 'textprocessor_openai'

        valid_indices: List[int] = []
        valid_prompts: List[str] = []
        valid_row_names: List[List[str]] = []
        valid_col_names: List[List[str]] = []
        for idx, item in enumerate(batch_items):
            values_matrix = item.get('values_matrix')
            if values_matrix is None:
                values_matrix = item.get('values', [])
            row_names = item.get('row_names') or item.get('index_names')
            column_names = item.get('column_names')

            values_matrix, row_names, column_names = self._normalize_matrix_inputs(
                values_matrix, row_names, column_names
            )
            if self._count_numeric_cells(values_matrix) < 2:
                continue
            full_prompt = (
                f"{prompt_template['system_prompt']}\n\n"
                f"{prompt_template['user_prompt'].format(values_matrix=values_matrix, row_names=row_names, column_names=column_names)}"
            )
            valid_indices.append(idx)
            valid_prompts.append(full_prompt)
            valid_row_names.append(row_names)
            valid_col_names.append(column_names)

        if not valid_prompts:
            return results

        m_print(f" Using TextProcessor /chat/batch for {len(valid_prompts)} windows")
        for start in range(0, len(valid_prompts), batch_size):
            end = min(start + batch_size, len(valid_prompts))
            chunk_prompts = valid_prompts[start:end]
            chunk_indices = valid_indices[start:end]
            chunk_rows = valid_row_names[start:end]
            chunk_cols = valid_col_names[start:end]

            payload = {
                'prompts': chunk_prompts,
                'provider': provider,
                'model': model,
                'max_tokens': 2000,
                'temperature': 0.1,
                'parallel': True
            }

            try:
                response = self.session.post(batch_url, json=payload, timeout=180)
                if response.status_code != 200:
                    m_print(f" /chat/batch error: {response.status_code} - {response.text}")
                    continue

                response.encoding = 'utf-8'
                data = response.json()
                batch_results = data.get('results', [])
                if not isinstance(batch_results, list) or not batch_results:
                    m_print(" /chat/batch response has no results")
                    continue

                for local_i, item in enumerate(batch_results):
                    idx_in_chunk = item.get('index', local_i)
                    if not isinstance(idx_in_chunk, int) or idx_in_chunk < 0 or idx_in_chunk >= len(chunk_indices):
                        continue
                    global_idx = chunk_indices[idx_in_chunk]
                    row_names = chunk_rows[idx_in_chunk]
                    col_names = chunk_cols[idx_in_chunk]

                    if item.get('error'):
                        m_print(f" /chat/batch item {global_idx} failed: {item.get('error')}")
                        continue

                    tp_result = item.get('result')
                    if not isinstance(tp_result, dict):
                        continue

                    content = tp_result.get('output') or tp_result.get('response') or tp_result.get('message')
                    if not content:
                        continue

                    parsed_result = self._parse_llm_response(str(content), llm_type)
                    if not parsed_result:
                        continue

                    parsed_result['_row_names'] = row_names
                    parsed_result['_column_names'] = col_names
                    parsed_result['_index_names'] = row_names  # backward-compatible alias
                    parsed_result['textprocessor_usage'] = tp_result.get('usage', {})
                    parsed_result['textprocessor_billing'] = tp_result.get('billing_summary', {})
                    parsed_result['textprocessor_token_summary'] = tp_result.get('token_summary', {})
                    parsed_result['textprocessor_post_id'] = tp_result.get('post_id', 'unknown')
                    parsed_result['textprocessor_timestamp'] = tp_result.get('timestamp', 'unknown')
                    parsed_result['textprocessor_provider'] = tp_result.get('provider', provider)
                    parsed_result['textprocessor_model_alias'] = tp_result.get('model_alias', model)
                    parsed_result['textprocessor_raw_response'] = tp_result

                    results[global_idx] = parsed_result

            except Exception as e:
                m_print(f" /chat/batch request failed: {e}")
                continue

        return results

    def _detect_with_textprocessor(
        self,
        values_matrix: List[List[Any]],
        row_names: List[str],
        column_names: List[str],
    ) -> Optional[Dict[str, Any]]:
        return _detect_with_local_llm(self, values_matrix, row_names, column_names)
    
    def _detect_with_openai(
        self,
        values_matrix: List[List[Any]],
        row_names: List[str],
        column_names: List[str],
    ) -> Optional[Dict[str, Any]]:
        """使用TextProcessor的OpenAI API偵測"""
        m_print(" 使用TextProcessor的OpenAI API偵測數學規律")
        
        try:
            # 準備提示
            prompt_template = self.prompts.get('math_rule')
            if not prompt_template:
                m_print(" 找不到數學規律提示模板")
                return None
            
            # 構建完整的用戶提示（包含系統提示）
            full_prompt = (
                f"{prompt_template['system_prompt']}\n\n"
                f"{prompt_template['user_prompt'].format(values_matrix=values_matrix, row_names=row_names, column_names=column_names)}"
            )
            
            # 呼叫TextProcessor的OpenAI API
            # 將模型名稱轉換為TextProcessor的別名格式
            model_alias = self.openai_model
            if model_alias == 'gpt-4o':
                model_alias = 'gpt4o_chat'
            elif model_alias == 'gpt-4':
                model_alias = 'gpt4_chat'
            elif model_alias == 'gpt-3.5-turbo':
                model_alias = 'gpt35_chat'  # 或其他可用的別名
            
            payload = {
                'prompt': full_prompt,
                'provider': 'openai',
                'model': model_alias,
                'max_tokens': 2000,
                'temperature': 0.1
            }
            
            m_print(f" 使用OpenAI模型別名: {model_alias}")
            
            response = self.session.post(
                self.textprocessor_url,
                json=payload,
                timeout=120
            )
            
            if response.status_code != 200:
                m_print(f" TextProcessor OpenAI API錯誤: {response.status_code} - {response.text}")
                return None
            
            try:
                # 確保回應編碼正確
                response.encoding = 'utf-8'
                result = response.json()
            except json.JSONDecodeError as e:
                m_print(f" TextProcessor OpenAI回應JSON解析失敗: {e}")
                m_print(f" 原始回應: {response.text}")
                return None
            
            # 檢查TextProcessor回應格式
            # m_print(f" TextProcessor OpenAI原始回應: {result}")
            
            # 檢查是否有錯誤
            if 'error' in result:
                m_print(f" TextProcessor OpenAI回傳錯誤: {result['error']}")
                return None
            
            # 提取回應內容 - TextProcessor的主要內容在'output'欄位
            content = result.get('output')
            if not content:
                m_print(" TextProcessor OpenAI回應為空")
                # 檢查是否有其他可能的內容欄位
                if 'response' in result:
                    content = result.get('response')
                    m_print(" 使用'response'欄位作為內容")
                elif 'message' in result:
                    content = result.get('message')
                    m_print(" 使用'message'欄位作為內容")
                else:
                    m_print(" 找不到任何內容欄位")
                    return None
            
            # 調試：檢查內容格式
            m_print(f" 提取的內容類型: {type(content)}")
            m_print(f" 提取的內容長度: {len(str(content))}")
            m_print(f" 內容前100字符: {str(content)[:100]}")
            
            # 確保內容是字串格式
            if not isinstance(content, str):
                content = str(content)
                m_print(" 將內容轉換為字串格式")
            
            # 記錄TextProcessor的元資訊
            provider = result.get('provider', 'unknown')
            model_alias = result.get('model_alias', 'unknown')
            post_id = result.get('post_id', 'unknown')
            timestamp = result.get('timestamp', 'unknown')
            
            m_print(f" TextProcessor OpenAI Provider: {provider}")
            m_print(f" TextProcessor OpenAI Model: {model_alias}")
            m_print(f" TextProcessor OpenAI Post ID: {post_id}")
            
            # 記錄計價資訊
            usage_info = result.get('usage', {})
            billing_info = result.get('billing_summary', {})
            token_info = result.get('token_summary', {})
            
            if usage_info:
                m_print(f" OpenAI Token使用量: {usage_info}")
            if billing_info:
                m_print(f" OpenAI 計價資訊: {billing_info}")
            if token_info:
                m_print(f" OpenAI Token摘要: {token_info}")
            
            # 解析回應
            parsed_result = self._parse_llm_response(content, 'textprocessor_openai')
            
            # 將TextProcessor的完整資訊添加到結果中
            if parsed_result:
                parsed_result['textprocessor_usage'] = usage_info
                parsed_result['textprocessor_billing'] = billing_info
                parsed_result['textprocessor_token_summary'] = token_info
                parsed_result['textprocessor_post_id'] = post_id
                parsed_result['textprocessor_timestamp'] = timestamp
                parsed_result['textprocessor_provider'] = provider
                parsed_result['textprocessor_model_alias'] = model_alias
                parsed_result['textprocessor_raw_response'] = result
            
            return parsed_result
            
        except Exception as e:
            m_print(f" TextProcessor OpenAI呼叫失敗: {e}")
            return None
    
    def _split_equation(self, equation: str):
        if not equation or '=' not in equation:
            return None
        left, right = equation.split('=', 1)
        return left.strip(), right.strip()

    def _normalize_index_tokens(self, expr: str) -> str:
        """Normalize index tokens into $(a,b) form."""
        import re
        s = str(expr)
        # Convert bare "(r,c)" into "$(r,c)" for compatibility
        s = re.sub(r'(?<!\$)\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)', r'$(\1,\2)', s)
        # Normalize legacy $i tokens into $(i,0)
        s = re.sub(r'\$(\d+)', r'$(\1,0)', s)
        return s

    def _normalize_equation_sides(self, rule: Dict[str, Any], name_to_index: Dict[str, int]) -> None:
        """
        讓每條規則都有可驗證的 equation_sides（$(a,b) 形式）
        優先：有效的 equation_sides -> 若無，從 equation 轉 -> 再不行，用欄位名映射替換
        """
        sides = rule.get('equation_sides')
        eq    = rule.get('equation')

        def _set_equation_from_sides():
            if isinstance(rule.get('equation_sides'), list) and len(rule['equation_sides']) == 2:
                rule['equation'] = f"{rule['equation_sides'][0]} = {rule['equation_sides'][1]}"

        # 1) 若 sides 合法，直接用
        if (isinstance(sides, list) and len(sides) == 2 and
            all(isinstance(x, str) for x in sides) and
            all(is_dollar_expr(x) for x in sides)):
            rule['equation_sides'] = [
                self._normalize_index_tokens(sides[0]),
                self._normalize_index_tokens(sides[1])
            ]
            _set_equation_from_sides()
            return

        # 2) 從 equation 回填 sides
        if isinstance(eq, str):
            sp = self._split_equation(eq)
            if sp and all(is_dollar_expr(x) for x in sp):
                rule['equation_sides'] = [
                    self._normalize_index_tokens(sp[0]),
                    self._normalize_index_tokens(sp[1]),
                ]
                _set_equation_from_sides()
                return

        # 3) 將中文欄位名替換成 $(row,0)（僅 row 名稱可推導）
        def replace_names(expr: str) -> str:
            s = str(expr)
            # 長名稱先替換，避免子字串誤替
            for name in sorted(name_to_index.keys(), key=len, reverse=True):
                s = s.replace(name, f"$({name_to_index[name]},0)")
            return s

        if isinstance(sides, list) and len(sides) == 2:
            left = replace_names(sides[0])
            right= replace_names(sides[1])
            if is_dollar_expr(left) and is_dollar_expr(right):
                rule['equation_sides'] = [
                    self._normalize_index_tokens(left),
                    self._normalize_index_tokens(right),
                ]
                _set_equation_from_sides()
                return

        # 4) 最後再試一次用 equation + 名稱替換
        if isinstance(eq, str):
            sp = self._split_equation(eq)
            if sp:
                left = replace_names(sp[0])
                right= replace_names(sp[1])
                if is_dollar_expr(left) and is_dollar_expr(right):
                    rule['equation_sides'] = [
                        self._normalize_index_tokens(left),
                        self._normalize_index_tokens(right),
                    ]
                    _set_equation_from_sides()
                    return

    def _parse_llm_response(self, content: str, llm_type: str) -> Optional[Dict[str, Any]]:
        """解析LLM回應"""
        try:
            # 清理回應內容
            content = content.strip()
            
            # 移除可能的markdown標記
            if content.startswith('```json'):
                content = content[7:]
            if content.endswith('```'):
                content = content[:-3]
            # 移除末尾的換行符和反引號
            content = content.rstrip('\n`').strip()
            
            # 調試：記錄要解析的內容
            # m_print(f" 準備解析JSON內容: {content[:200]}{'...' if len(content) > 200 else ''}")
            
            # 檢查內容是否包含特殊字符
            # if any(ord(c) > 127 for c in content[:100]):
            #     m_print(" 內容包含非ASCII字符")
            
            # 解析JSON
            try:
                result = json.loads(content)
            except json.JSONDecodeError as e:
                m_print(f" JSON解析失敗", colora=LOGger.FAIL)
                m_print(f" 原始內容: {content}", colora=LOGger.FAIL)
                
                # 嘗試修復常見的JSON問題
                # 1. 移除多餘的換行符和空白
                cleaned_content = content.replace('\n', '').replace('\r', '').strip()
                try:
                    result = json.loads(cleaned_content)
                    m_print(" 修復後JSON解析成功")
                except json.JSONDecodeError:
                    # 1.5. 嘗試處理轉義的換行符
                    try:
                        # 將 \n 轉換為實際的換行符
                        unescaped_content = content.replace('\\n', '\n').replace('\\r', '\r')
                        result = json.loads(unescaped_content)
                        m_print(" 轉義字符修復後JSON解析成功")
                    except json.JSONDecodeError:
                        # 2. 嘗試處理Unicode字符
                        try:
                            # 確保UTF-8編碼
                            if isinstance(content, str):
                                content_bytes = content.encode('utf-8')
                                content = content_bytes.decode('utf-8')
                            result = json.loads(content)
                            m_print(" Unicode修復後JSON解析成功")
                        except json.JSONDecodeError:
                            # 3. 提取純JSON部分（移除額外說明文字）
                            try:
                                # 找到第一個 '{' 和最後一個 '}'
                                s = content.find('{')
                                e = content.rfind('}')
                                if s != -1 and e != -1 and e > s:
                                    json_part = content[s:e+1]
                                    m_print(f" 提取JSON部分: {json_part[:100]}...")
                                    result = json.loads(json_part)
                                    m_print(" 提取JSON片段解析成功")
                                else:
                                    raise
                            except json.JSONDecodeError:
                                # 4. 最後嘗試：使用正則表達式提取JSON
                                try:
                                    import re
                                    # 尋找完整的JSON物件
                                    json_pattern = r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}'
                                    matches = re.findall(json_pattern, content)
                                    if matches:
                                        # 取最長的匹配
                                        json_part = max(matches, key=len)
                                        m_print(f" 正則提取JSON: {json_part[:100]}...")
                                        result = json.loads(json_part)
                                        m_print(" 正則提取JSON解析成功")
                                    else:
                                        raise
                                except json.JSONDecodeError:
                                    raise

            # 允許呼叫端在 result 中塞入 index_names 方便映射（見 B）
            name_to_index = {}
            if isinstance(result.get('_index_names'), list):
                for i, nm in enumerate(result['_index_names']):
                    if nm is not None:
                        name_to_index[str(nm)] = i

            # 規則回填/正規化
            if 'rules' in result and isinstance(result['rules'], list):
                for r in result['rules']:
                    self._normalize_equation_sides(r, name_to_index)
            
            # 添加元資訊
            result['llm_type'] = llm_type
            result['raw_response'] = content
            
            # 驗證結果格式
            if not self._validate_result(result):
                m_print(" LLM回應格式不正確")
                m_print(f" 回應內容: {json.dumps(result, indent=2, ensure_ascii=False)[:500]}...")
                return None
            
            # m_print(f" {llm_type}偵測成功")
            return result
            
        except json.JSONDecodeError as e:
            m_print(f" JSON解析失敗: {e}")
            m_print(f"原始回應: {content[:500]}...")  # 只顯示前500字符
            return None
        except Exception as e:
            m_print(f" 回應解析失敗: {e}")
            return None
    
    def _validate_result(self, result: Dict[str, Any]) -> bool:
        """驗證結果格式"""
        # 必要欄位，confidence可以補預設值
        if 'rules' not in result:
            m_print(" 缺少必要欄位: rules")
            return False
        
        # 如果缺少description，補預設值
        if 'description' not in result:
            result['description'] = '偵測到的數學規律'
        
        # 如果缺少confidence，補預設值
        if 'confidence' not in result:
            result['confidence'] = 0.8
        
        if not isinstance(result['rules'], list):
            m_print(" rules必須是列表")
            return False
        
        for i, rule in enumerate(result['rules']):
            if not isinstance(rule, dict):
                m_print(" 每個rule必須是字典")
                return False
            
            # 必要欄位檢查 - 新格式只需要equation_sides
            if 'equation_sides' not in rule:
                m_print(f" rule {i} 缺少equation_sides欄位")
                self._coerce_equation_sides_from_equation(rule)
                if 'equation_sides' not in rule:
                    m_print(f" rule 補正後仍缺少equation_sides欄位，無法驗證")
                    return False
            
            # 檢查equation_sides格式
            equation_sides = rule.get('equation_sides', [])
            if not isinstance(equation_sides, list) or len(equation_sides) != 2:
                m_print(f" rule {i} equation_sides必須是包含2個元素的列表")
                return False
            
            # 檢查confidence欄位
            if 'confidence' not in rule:
                m_print(f" rule {i} 缺少confidence欄位")
                # 提供預設值
                rule['confidence'] = 0.5
        
        return True
    
    def _verify_equation_sides(self, equation_sides: List[str], values: List[List[Any]]) -> bool:
        """
        驗證equation_sides格式的數學正確性
        
        Args:
            equation_sides: 等式兩邊，如 ["$(0,0) + $(0,1)", "$(1,0)"]
            values: 數值矩陣（rows x cols）
            
        Returns:
            bool: 是否數學正確
        """
        try:
            if len(equation_sides) != 2:
                return False

            if not self._equation_sides_in_range(equation_sides, values):
                m_print(f"equation_sides索引超出範圍，略過驗證: {equation_sides}")
                return False
            
            left_expr, right_expr = equation_sides
            
            # 計算左邊表達式的值
            left_value = self._evaluate_dollar_expression(left_expr, values)
            # 計算右邊表達式的值
            right_value = self._evaluate_dollar_expression(right_expr, values)
            
            # 檢查是否相等（容差0.01）
            tolerance = 0.01
            return abs(left_value - right_value) < tolerance
            
        except Exception as e:
            m_print(f"評估equation_sides失敗: {equation_sides}, 錯誤: {e}")
            return False

    def _equation_sides_in_range(self, equation_sides: List[str], values: List[List[Any]]) -> bool:
        """
        檢查 equation_sides 中所有索引是否在 values 的範圍內。
        """
        import re

        def _to_matrix(v):
            if not isinstance(v, list):
                return []
            if v and not isinstance(v[0], list):
                return [[x] for x in v]
            return v

        matrix = _to_matrix(values)
        max_rows = len(matrix)
        max_cols = max((len(r) for r in matrix), default=0)

        def _check_expr(expr: str) -> bool:
            s = re.sub(r'(?<!\$)\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)', r'$(\1,\2)', str(expr))
            found = False
            for r_str, c_str in re.findall(r'\$\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)', s):
                found = True
                r = int(r_str)
                c = int(c_str)
                if r < 0 or c < 0 or r >= max_rows or c >= max_cols:
                    return False
            for r_str in re.findall(r'\$(\d+)', s):
                found = True
                r = int(r_str)
                if r < 0 or r >= max_rows or 0 >= max_cols:
                    return False
            return found

        if not equation_sides or len(equation_sides) != 2:
            return False
        return _check_expr(equation_sides[0]) and _check_expr(equation_sides[1])
    
    def _evaluate_dollar_expression(self, expression: str, values: List[List[Any]]) -> float:
        """
        評估$符號表達式的值
        
        Args:
            expression: 如 "$(0,0) + $(1,0)" 或 "$(0,1)"
            values: 數值矩陣（rows x cols）
            
        Returns:
            float: 計算結果
        """
        import re
        
        def _to_matrix(v):
            if not isinstance(v, list):
                return []
            if v and not isinstance(v[0], list):
                return [[x] for x in v]
            return v

        matrix = _to_matrix(values)
        max_rows = len(matrix)
        max_cols = max((len(r) for r in matrix), default=0)
        matrix = [list(r) + [None] * (max_cols - len(r)) for r in matrix]

        def _get_value(r: int, c: int) -> float:
            if r < 0 or r >= max_rows or c < 0 or c >= max_cols:
                raise ValueError(f"索引超出範圍: ({r},{c})")
            val = matrix[r][c]
            if val is None:
                raise ValueError(f"索引 ({r},{c}) 對應值為空")
            try:
                return float(val)
            except Exception:
                raise ValueError(f"索引 ({r},{c}) 對應值無法轉為數值: {val}")

        # Normalize bare "(r,c)" into "$(r,c)" before evaluation
        expression = re.sub(r'(?<!\$)\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)', r'$(\1,\2)', str(expression))

        # 先替換 $(a,b)
        def replace_2d(match):
            r = int(match.group(1))
            c = int(match.group(2))
            return str(_get_value(r, c))

        numeric_expression = re.sub(r'\$\(\s*(\d+)\s*,\s*(\d+)\s*\)', replace_2d, expression)

        # 相容舊式 $i（視為 $(i,0)）
        def replace_1d(match):
            idx = int(match.group(1))
            return str(_get_value(idx, 0))

        numeric_expression = re.sub(r'\$(\d+)', replace_1d, numeric_expression)
        
        # 安全地評估數學表達式
        return self._safe_eval_math_expression_simple(numeric_expression)
    
    
    def _safe_eval_math_expression_simple(self, expression: str) -> float:
        """
        安全地評估數學表達式（簡化版）
        
        Args:
            expression: 數學表達式，如 "237.6 + 1.2"
            
        Returns:
            float: 計算結果
        """
        import re
        
        # 只允許數字、小數點、基本運算符、括號和空格
        if not re.match(r'^[\d\.\+\-\*\/\(\)\s]+$', expression):
            raise ValueError(f"表達式包含不安全的字符: {expression}")
        
        # 使用eval計算（已經過安全檢查）
        try:
            result = eval(expression, {"__builtins__": {}}, {})
            return float(result)
        except Exception as e:
            raise ValueError(f"無法計算表達式: {expression}, 錯誤: {str(e)}")
    
    def validate_rules(self, values: List[List[Any]], rules: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        驗證偵測到的規律是否正確，支援新的math_expression格式
        
        Args:
            values: 原始數值矩陣
            rules: 偵測到的規律列表
            
        Returns:
            Dict: 驗證結果
        """
        validation_results = []
        valid_count = 0
        
        for rule in rules:
            rule_text = rule.get('equation', rule.get('rule', rule.get('description', '')))
            equation_sides = rule.get('equation_sides', [])  # 新增equation_sides支援
            
            try:
                # 簡化驗證 - 只檢查equation_sides的數學正確性
                math_valid = False
                math_error = None
                
                # 使用equation_sides進行驗證
                if equation_sides and len(equation_sides) == 2:
                    try:
                        math_valid = self._verify_equation_sides(equation_sides, values)
                        if not math_valid:
                            math_error = f"equation_sides驗證失敗: {equation_sides}"
                    except Exception as e:
                        math_error = f"equation_sides驗證異常: {str(e)}"
                        math_valid = False
                else:
                    math_error = "equation_sides格式不正確或缺失"
                
                # 簡化判斷：只需要數學計算正確
                is_valid = math_valid
                
                validation_result = {
                    'rule': rule_text,
                    'is_valid': is_valid,
                    'validation_method': 'equation_sides',
                    'math_valid': math_valid
                }
                
                # 添加數學驗證詳情
                if equation_sides and len(equation_sides) == 2:
                    # 左右數值(若能計算就附上)
                    left_val = right_val = None
                    try:
                        left_val  = self._evaluate_dollar_expression(equation_sides[0], values)
                        right_val = self._evaluate_dollar_expression(equation_sides[1], values)
                    except Exception:
                        pass

                    validation_result['equation_preview'] = f"{equation_sides[0]} = {equation_sides[1]}"
                    validation_result['math_verification'] = {
                        'equation_sides': equation_sides,
                        'verification_type': 'equation_sides',
                        'is_mathematically_correct': math_valid,
                        'left_value': left_val,
                        'right_value': right_val,
                        'error': math_error
                    }
                
                validation_results.append(validation_result)
                
                # 調試信息
                if not is_valid:
                    m_print(f"   規則驗證失敗: {rule_text[:50]}...")
                    m_print(f"   數學有效: {math_valid}")
                    if math_error:
                        m_print(f"   數學錯誤: {math_error}")
                
                if is_valid:
                    valid_count += 1
                    
            except Exception as e:
                validation_results.append({
                    'rule': rule_text,
                    'is_valid': False,
                    'validation_method': 'error',
                    'error': f'驗證錯誤: {e}'
                })
        
        total_rules = len(rules)
        success_rate = (valid_count / total_rules) if total_rules > 0 else 0
        
        return {
            'total_rules': total_rules,
            'valid_rules': valid_count,
            'success_rate': success_rate,
            'validation_details': validation_results
        }
    
    def _verify_equation_mathematically(self, equation_with_values):
        """
        數學驗證等式是否正確
        
        Args:
            equation_with_values: 包含數值的等式，如 "7.0 = 2.5 - 224.4"
            
        Returns:
            Tuple[bool, float, str]: (是否正確, 計算結果, 錯誤信息)
        """
        try:
            # 分割等式的左右兩邊
            if '=' not in equation_with_values:
                return False, None, "等式中沒有等號"
            
            parts = equation_with_values.split('=')
            if len(parts) != 2:
                return False, None, "等式格式不正確"
            
            left_side = parts[0].strip()
            right_side = parts[1].strip()
            
            # 嘗試計算左邊的值（可能是數值或表達式）
            try:
                # 先嘗試直接解析為數值
                left_value = float(left_side)
            except ValueError:
                # 如果不是純數值，嘗試作為表達式計算
                try:
                    left_value = self._safe_eval_math_expression(left_side)
                except Exception as e:
                    return False, None, f"無法計算左邊的表達式: {left_side}, 錯誤: {str(e)}"
            
            # 嘗試計算右邊的值（可能是數值、表達式或欄位名）
            try:
                # 先嘗試直接解析為數值
                right_value = float(right_side)
            except ValueError:
                # 如果不是純數值，檢查是否包含中文字符（欄位名）
                if self._contains_chinese_or_field_names(right_side):
                    # 如果右邊是欄位名，我們無法直接驗證，返回警告但不標記為錯誤
                    return True, None, f"右邊包含欄位名稱，無法直接驗證: {right_side}"
                else:
                    # 嘗試作為數學表達式計算
                    try:
                        right_value = self._safe_eval_math_expression(right_side)
                    except Exception as e:
                        return False, None, f"無法計算右邊的表達式: {right_side}, 錯誤: {str(e)}"
            
            # 檢查兩邊是否相等（允許小的浮點誤差）
            tolerance = 0.01
            difference = abs(left_value - right_value)
            is_correct = difference < tolerance
            
            error_msg = None if is_correct else f"等式不成立: {left_value} ≠ {right_value} (差異: {difference:.6f})"
            
            return is_correct, right_value, error_msg
            
        except Exception as e:
            return False, None, f"數學驗證異常: {str(e)}"
    
    def _safe_eval_math_expression(self, expression):
        """
        安全地評估數學表達式
        
        Args:
            expression: 數學表達式字符串
            
        Returns:
            float: 計算結果
        """
        import re
        
        # 只允許數字、小數點、基本運算符和空格
        if not re.match(r'^[\d\.\+\-\*\/\(\)\s]+$', expression):
            raise ValueError(f"表達式包含不安全的字符: {expression}")
        
        # 使用eval計算（已經過安全檢查）
        try:
            result = eval(expression)
            return float(result)
        except Exception as e:
            raise ValueError(f"無法計算表達式: {expression}, 錯誤: {str(e)}")
    
    def _contains_chinese_or_field_names(self, text):
        """
        檢查文本是否包含中文字符或欄位名稱
        
        Args:
            text: 要檢查的文本
            
        Returns:
            bool: 是否包含中文或欄位名稱
        """
        import re
        
        # 檢查是否包含中文字符
        chinese_pattern = r'[\u4e00-\u9fff]'
        if re.search(chinese_pattern, text):
            return True
        
        # 檢查是否包含常見的欄位關鍵詞
        field_keywords = ['重', '量', '後', '前', '電鍍', '磨光', '研磨', 'AC', 'NPS', '銘板', '膠套']
        for keyword in field_keywords:
            if keyword in text:
                return True
        
        return False

    def _coerce_equation_sides_from_equation(self, rule: Dict[str, Any]) -> None:
        """
        若缺少 equation_sides、但 rule['equation'] 形如 "$(0,0) + $(0,1) = $(1,0)"，則自動補上：
        equation_sides = ["$(0,0) + $(0,1)", "$(1,0)"]
        """
        eq = (rule.get('equation') or rule.get('equation_with_indices') or '').strip()
        if not eq or '$' not in eq or '=' not in eq:
            return
        left, right = eq.split('=', 1)
        left, right = left.strip(), right.strip()
        if left and right and 'equation_sides' not in rule:
            rule['equation_sides'] = [left, right]

def main():
    """主函數"""
    parser = argparse.ArgumentParser(description='LLM數學規律偵測器')
    parser.add_argument('--values', type=float, nargs='+', 
                       help='數值列表，例如: --values 1.2 2.4 3.6')
    parser.add_argument('--columns', type=str, nargs='+', 
                       help='對應的欄位名稱')
    parser.add_argument('--use-openai', action='store_true', 
                       help='強制使用TextProcessor OpenAI API（預設使用TextProcessor 8B地端核心）')
    parser.add_argument('--test', action='store_true', 
                       help='使用測試資料')
    
    args = parser.parse_args()
    
    # 初始化偵測器
    prefer_local = not args.use_openai
    detector = LLMMathRuleDetector(prefer_local=prefer_local)
    
    # 準備測試資料
    if args.test:
        test_values = [238.8, 237.6, 238.8, 1.2, 237.6]
        test_columns = ["index_1", "index_2", "index_3", "index_4", "index_5"]
        print("使用測試資料:")
        print(f"Values: {test_values}")
        print(f"Columns: {test_columns}")
    elif args.values:
        test_values = args.values
        test_columns = args.columns
    else:
        print("請提供 --values 參數或使用 --test 參數")
        return
    
    # 執行偵測
    values_matrix = [[v] for v in test_values]
    row_names = test_columns
    col_names = ["value"]
    result = detector.detect_math_rules(values_matrix, row_names, col_names, args.use_openai)
    
    if result:
        print("\n=== 偵測結果 ===")
        print(f"LLM類型: {result.get('llm_type', 'unknown')}")
        print(f"描述: {result.get('description', 'N/A')}")
        print(f"信心度: {result.get('confidence', 'N/A')}")
        
        print(f"\n發現 {len(result.get('rules', []))} 個規律:")
        for i, rule in enumerate(result.get('rules', []), 1):
            print(f"\n規律 {i}:")
            # 顯示equation_sides（新格式）
            if 'equation_sides' in rule:
                equation_sides = rule.get('equation_sides', [])
                if len(equation_sides) == 2:
                    print(f"  等式兩邊: {equation_sides[0]} = {equation_sides[1]}")
                else:
                    print(f"  等式兩邊: {equation_sides}")
            
            print(f"  信心度: {rule.get('confidence', 'N/A')}")
            
            # 顯示描述（如果有）
            if 'description' in rule:
                print(f"  描述: {rule.get('description', 'N/A')}")
        
        # 驗證規律
        validation = detector.validate_rules(values_matrix, result.get('rules', []))
        print(f"\n=== 驗證結果 ===")
        print(f"總規律數: {validation['total_rules']}")
        print(f"有效規律數: {validation['valid_rules']}")
        print(f"成功率: {validation['success_rate']:.2%}")
        
    else:
        print(" 偵測失敗")

if __name__ == "__main__":
    main()
