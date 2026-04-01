#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
數學規則偵測器
使用語言模型來記錄和分析Excel數據中的潛在數學規則
"""

import os
import sys
import json
import argparse
import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime
from dotenv import load_dotenv
import openai
import requests
import re

from utils import LOGger

m_logfile = Path(__file__).resolve().parent.parent / "outputs" / "math_rule_detector.log"
# 載入環境變數
load_dotenv()

class MathRule:
    def __init__(self, ID: str, eval_string: str):
        self.ID = ID
        self.eval_string = eval_string

    def execute(self, *args, **kwargs):
        return eval(self.eval_string)(*args, **kwargs)


    

class MathRuleDetector:
    """數學規則偵測器"""
    
    def __init__(self, window_shape: tuple = (5, 1)):
        """初始化偵測器
        
        Args:
            window_shape: 窗格形狀 (height, width)，預設為 (5, 1)
        """
        self.window_shape = window_shape
        self.window_size = window_shape[0]  # 取高度作為分析的數值個數
        self.llm_api_url = os.getenv('LLM_API_URL', 'http://localhost:11434/v1').rstrip('/')
        self.llm_model_name = os.getenv('LLM_MODEL_NAME', 'llama3.1')
        self.api_key = os.getenv('LLM_API_KEY')
        self.openai_api_key = os.getenv('OPENAI_API_KEY')
        
        # 初始化客戶端
        self.openai_client = None
        self.session = None
        
        if self.openai_api_key and self.openai_api_key != 'your_openai_api_key_here':
            self.openai_client = openai.OpenAI(api_key=self.openai_api_key)
            print("OpenAI客戶端已初始化")
        else:
            print("OpenAI API金鑰未設定，將使用本地LLM")
        
        # 建立 session（參考check_llm_connection.py）
        self.session = requests.Session()
        
        # 設定基本標頭（完全按照check_llm_connection.py）
        base_headers = {
            'Content-Type': 'application/json',
            'User-Agent': 'LLM-API-Client/1.0'  # 使用與check_llm_connection.py相同的User-Agent
        }
        self.session.headers.update(base_headers)
        
        # 初始化日誌
        # self.logger = LOGger.addloger(logfile='')
        self.logger = LOGger.addloger(logfile=m_logfile)
        
        # 測試地端LLM連接（如果沒有使用OpenAI）
        if not self.openai_client:
            self.logger("測試地端LLM連接...")
            try:
                if self.test_local_llm_connection():
                    self.logger("地端LLM連接正常")
                else:
                    self.logger("警告: 地端LLM連接測試失敗，但不影響正常使用")
            except Exception as e:
                self.logger(f"地端LLM連接測試異常: {e}，將繼續嘗試使用")
        
    def load_excel_data(self, file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """載入Excel數據"""
        try:
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)
            
            self.logger(f"成功載入Excel文件: {file_path}")
            if sheet_name:
                self.logger(f"工作表: {sheet_name}")
            self.logger(f"數據形狀: {df.shape}")
            
            return df
        except Exception as e:
            self.logger(f"載入Excel文件失敗: {e}")
            raise
    
    def extract_numeric_patterns(self, df: pd.DataFrame, start_loc_row_col: str) -> Dict[str, List[float]]:
        """提取數值模式"""
        patterns = {}
        
        # 先檢查可用的欄位
        self.logger(f"可用欄位: {list(df.columns)}")
        
        # 根據start_loc_row_name分組
        if 'start_loc_row_name' in df.columns:
            grouped = df.groupby('start_loc_row_name')
            
            for name, group in grouped:
                # 提取數值 - 根據window_size動態調整
                numeric_cols = []
                for i in range(1, self.window_size + 1):  # value_1 到 value_N
                    col_name = f'value_{i}'
                    if col_name in group.columns:
                        values = group[col_name].dropna().tolist()
                        if values:
                            value = values[0]  # 取第一個值
                            # 確保是數值類型
                            try:
                                if isinstance(value, str):
                                    # 嘗試轉換字串為數字
                                    numeric_value = float(value.replace(',', '').strip())
                                else:
                                    numeric_value = float(value)
                                numeric_cols.append(numeric_value)
                                self.logger(f"成功轉換 {col_name}: {value} -> {numeric_value}")
                            except (ValueError, TypeError):
                                self.logger(f"無法轉換為數值 {col_name}: `{value}` (類型: {type(value)})")
                                # 跳過非數值項目
                                continue
                    else:
                        self.logger(f"欄位 {col_name} 不存在")
                
                if len(numeric_cols) >= self.window_size:
                    patterns[f"start_loc_row_name_{name}"] = numeric_cols[:self.window_size]
                    self.logger(f"成功提取模式 {name}: {numeric_cols[:self.window_size]}")
                else:
                    self.logger(f"模式 {name} 數值不足{self.window_size}個 (只有{len(numeric_cols)}個)，跳過")
        
        return patterns
    
    def load_prompt_template(self) -> Dict[str, str]:
        """載入提示詞模板"""
        try:
            prompt_file = os.path.join(os.path.dirname(__file__), '..', 'prompts', 'math_rule_detection.json')
            with open(prompt_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            self.logger(f"載入提示詞文件失敗: {e}")
            # 提供備用提示詞
            return {
                "system_prompt": "你是一個專業的數學分析專家，專門找出數值之間的線性關係。請務必只回傳有效的JSON格式。",
                "user_prompt": "分析數值: {values}\n對應欄位: {column_names}\n\n找出線性關係，係數只能是 -1, 0, 1\n\n回傳JSON格式：\n{{\n  \"rules\": [\n    {{\n      \"coefficients\": [a1, a2, a3, a4, a5],\n      \"equation\": \"數學式\",\n      \"confidence\": 0.95\n    }}\n  ],\n  \"description\": \"簡短描述\"\n}}"
            }

    def generate_rule_prompt(self, values: List[float], column_names: List[str]) -> str:
        """生成規則識別提示詞"""
        template = self.load_prompt_template()
        
        # 使用模板生成提示詞
        user_prompt = template.get("user_prompt", "")
        prompt = user_prompt.format(
            values=values, 
            column_names=column_names
        )
        
        self.logger(f"使用提示詞模板生成提示")
        return prompt
    
    def test_local_llm_connection(self) -> bool:
        """測試地端LLM連接（參考check_llm_connection.py）"""
        try:
            # 檢查基本連線
            base_url = self.llm_api_url.replace('/v1', '')
            self.logger(f"測試連接: {base_url}/api/tags")
            
            response = requests.get(f"{base_url}/api/tags", timeout=3)
            
            if response.status_code == 200:
                models = response.json().get('models', [])
                model_exists = any(self.llm_model_name in m['name'] for m in models)
                self.logger(f"可用模型: {[m['name'] for m in models]}")
                self.logger(f"指定模型 '{self.llm_model_name}' 是否存在: {model_exists}")
                return model_exists
            else:
                self.logger(f"基本連線失敗: {response.status_code}")
                return False
                
        except requests.exceptions.Timeout:
            self.logger(f"連接超時，但LLM服務可能仍可用")
            return True  # 超時不代表服務不可用，返回True繼續嘗試
        except requests.exceptions.ConnectionError:
            self.logger(f"無法連接到LLM服務: {base_url}")
            return False
        except Exception as e:
            self.logger(f"LLM連接測試異常: {e}")
            return True  # 其他異常也不影響繼續嘗試

    def call_llm(self, prompt: str, use_openai: bool = False, values: List[float] = None) -> str:
        """呼叫語言模型（參考check_llm_connection.py的方法）"""
        try:
            # 載入系統提示詞
            template = self.load_prompt_template()
            system_prompt = template.get("system_prompt", "你是一個數學分析專家，專門找出數值之間的線性關係。請只回傳有效的JSON格式。")
            
            # 動態添加索引範圍限制到system_prompt
            if values:
                max_index = len(values) - 1
                available_indices = ", ".join([f"${i}" for i in range(len(values))])
                
                system_prompt += f"\n\n**當前數據限制：**\n"
                system_prompt += f"- 可用的索引範圍：{available_indices}\n"
                system_prompt += f"- 最大索引為 ${max_index}，請勿使用超出範圍的索引\n"
                system_prompt += f"- 數據共有 {len(values)} 個值，索引從 $0 到 ${max_index}"
            
            if use_openai and self.openai_client:
                response = self.openai_client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.1,
                    max_tokens=1000
                )
                return response.choices[0].message.content.strip()
            else:
                # 使用地端LLM（參考check_llm_connection.py的方法）
                messages = [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt}
                ]
                
                # 使用與check_llm_connection.py相同的payload格式
                payload = {
                    'model': self.llm_model_name,
                    'messages': messages,
                    'temperature': 0.1,
                    'max_tokens': 1000
                }
                
                # 準備請求標頭（本地LLM通常不需要API金鑰）
                headers = {}
                if self.api_key:
                    headers['Authorization'] = f'Bearer {self.api_key}'
                    self.logger(f"使用API金鑰進行認證")
                else:
                    self.logger(f"本地LLM不使用API金鑰")
                
                self.logger(f"呼叫地端LLM: {self.llm_api_url}")
                self.logger(f"使用模型: {self.llm_model_name}")
                
                response = self.session.post(
                    self.llm_api_url,  # 直接使用配置的URL
                    json=payload,
                    headers=headers,
                    timeout=30  # 縮短超時時間，避免長時間等待
                )
                
                self.logger(f"LLM API狀態: {response.status_code}")
                
                if response.status_code == 200:
                    result = response.json()
                    
                    # 處理不同的回應格式（參考check_llm_connection.py邏輯）
                    if 'choices' in result and len(result['choices']) > 0:
                        content = result['choices'][0]['message']['content']
                        self.logger(f"LLM回應成功，內容長度: {len(content)}")
                        return content.strip()
                    elif 'response' in result:
                        self.logger(f"LLM回應成功（response格式），內容長度: {len(result['response'])}")
                        return result['response'].strip()
                    elif 'error' in result:
                        raise Exception(f"LLM回應錯誤: {result['error']}")
                    else:
                        raise Exception(f"未知回應格式: {result}")
                else:
                    raise Exception(f"LLM API錯誤: {response.status_code} - {response.text}")
                    
        except requests.exceptions.Timeout as e:
            self.logger(f"LLM請求超時: {e}")
            self.logger("建議: LLM服務可能正在處理其他請求，請稍後再試或使用 --use-openai 參數")
            raise
        except requests.exceptions.ConnectionError as e:
            self.logger(f"無法連接到LLM服務: {e}")
            self.logger("建議: 檢查LLM服務是否運行，或嘗試使用 --use-openai 參數")
            raise
        except Exception as e:
            self.logger(f"LLM呼叫失敗: {e}")
            # 提供更詳細的錯誤信息和建議
            if "404" in str(e):
                self.logger("建議: 檢查LLM服務是否運行，或嘗試使用 --use-openai 參數")
            elif "Connection" in str(e):
                self.logger("建議: 檢查LLM服務地址和端口設定")
            raise
    
    def parse_llm_response(self, response: str) -> Dict[str, Any]:
        """解析LLM回應"""
        try:
            # 首先嘗試直接解析整個回應
            try:
                return json.loads(response)
            except json.JSONDecodeError:
                pass
            
            # 嘗試提取JSON陣列部分 [...]，使用非貪婪匹配
            array_match = re.search(r'\[[\s\S]*?\](?=\s*$|\s*\n\s*[^\]\}])', response)
            if not array_match:
                # 如果沒找到，嘗試更寬鬆的匹配
                array_match = re.search(r'\[[\s\S]*\]', response)
            
            if array_match:
                json_str = array_match.group().strip()
                self.logger(f"提取的JSON陣列: {json_str[:200]}...")
                
                try:
                    rules = json.loads(json_str)
                    return {"rules": rules, "description": "從陣列格式解析"}
                except json.JSONDecodeError as e:
                    self.logger(f"標準JSON解析失敗: {e}")
                    # 嘗試修復常見問題
                    fixed_json = json_str
                    
                    # 修復布林值
                    fixed_json = fixed_json.replace('true', 'True').replace('false', 'False')
                    
                    # 修復可能的尾隨逗號
                    fixed_json = re.sub(r',(\s*[\]\}])', r'\1', fixed_json)
                    
                    try:
                        rules = eval(fixed_json)  # 使用eval解析Python格式
                        # 轉換回標準格式
                        if isinstance(rules, list):
                            for rule in rules:
                                if isinstance(rule, dict):
                                    # 移除LLM提供的is_valid，我們會自己計算
                                    if 'is_valid' in rule:
                                        del rule['is_valid']
                            return {"rules": rules, "description": "從修復的陣列格式解析"}
                    except Exception as eval_error:
                        self.logger(f"eval解析也失敗: {eval_error}")
                        pass
            
            # 嘗試提取JSON物件部分 {...}
            json_match = re.search(r'\{.*?\}', response, re.DOTALL)
            if json_match:
                json_str = json_match.group()
                return json.loads(json_str)
            
            # 如果都失敗，嘗試手動解析係數
            self.logger("嘗試手動解析係數...")
            rules = []
            
            # 尋找係數模式 [1, -1, 1, -1, 1] 等
            coeff_pattern = r'\[([0-9\-,\s]+)\]'
            coeff_matches = re.findall(coeff_pattern, response)
            
            for i, match in enumerate(coeff_matches):
                try:
                    coeffs = [int(x.strip()) for x in match.split(',')]
                    if len(coeffs) == self.window_size:  # 確保係數數量正確
                        rules.append({
                            "coefficients": coeffs,
                            "equation": f"手動解析規則 {i+1}",
                            "confidence": 0.8  # 手動解析的信心度較低
                        })
                except:
                    continue
            
            if rules:
                return {"rules": rules, "description": "手動解析係數"}
            
            return {"rules": [], "description": "解析失敗"}
            
        except Exception as e:
            self.logger(f"JSON解析失敗: {e}")
            self.logger(f"原始回應前500字符: {response[:500]}")
            return {"rules": [], "description": "解析失敗"}
    
    def verify_rule(self, values: List[float], coefficients: List[float], tolerance: float = 0.01) -> Tuple[bool, float]:
        """驗證數學規則"""
        try:
            if len(values) != len(coefficients):
                self.logger(f"長度不匹配: values={len(values)}, coefficients={len(coefficients)}")
                return False, float('inf')
            
            # 確保所有值都是數值類型
            numeric_values = []
            for i, v in enumerate(values):
                try:
                    if isinstance(v, str):
                        numeric_v = float(v.replace(',', '').strip())
                    else:
                        numeric_v = float(v)
                    numeric_values.append(numeric_v)
                except (ValueError, TypeError):
                    self.logger(f"無法轉換值 {i}: {v} (類型: {type(v)})")
                    return False, float('inf')
            
            # 確保所有係數都是數值類型
            numeric_coefficients = []
            for i, c in enumerate(coefficients):
                try:
                    numeric_c = float(c)
                    numeric_coefficients.append(numeric_c)
                except (ValueError, TypeError):
                    self.logger(f"無法轉換係數 {i}: {c} (類型: {type(c)})")
                    return False, float('inf')
            
            # 計算線性組合
            result = sum(v * c for v, c in zip(numeric_values, numeric_coefficients))
            is_valid = abs(result) < tolerance
            
            self.logger(f"規則驗證: 值={numeric_values}, 係數={numeric_coefficients}, 結果={result}, 有效={is_valid}")
            
            return is_valid, result
        except Exception as e:
            self.logger(f"規則驗證失敗: {e}")
            self.logger(f"values類型: {[type(v) for v in values]}")
            self.logger(f"coefficients類型: {[type(c) for c in coefficients]}")
            return False, float('inf')
    
    def analyze_patterns(self, file_path: str, sheet_name: Optional[str] = None, use_openai: bool = False) -> Dict[str, Any]:
        """分析數據模式"""
        # 載入數據
        df = self.load_excel_data(file_path, sheet_name)
        
        # 提取數值模式
        patterns = self.extract_numeric_patterns(df, 'start_loc_row_name')
        
        results = {
            "program_path": os.path.abspath(__file__),
            "execution_time": datetime.now().isoformat(),
            "analyzed_file": os.path.abspath(file_path),
            "sheet_name": sheet_name,
            "total_patterns": len(patterns),
            "valid_rules": 0,
            "rule_compliance_ratio": 0.0,
            "patterns": {},
            "prompt_used": ""
        }
        
        column_names = [f"value_{i}" for i in range(1, self.window_size + 1)]
        
        for pattern_name, values in patterns.items():
            self.logger(f"分析模式: {pattern_name}")
            self.logger(f"數值: {values}")
            
            # 生成提示詞
            prompt = self.generate_rule_prompt(values, column_names)
            results["prompt_used"] = prompt
            
            try:
                # 呼叫LLM，傳入values以便動態添加索引限制
                llm_response = self.call_llm(prompt, use_openai, values)
                
                # 解析回應
                parsed_response = self.parse_llm_response(llm_response)
                
                # 驗證規則
                pattern_results = {
                    "values": values,
                    "llm_response": parsed_response,
                    "verified_rules": [],
                    "valid_count": 0
                }
                
                if "rules" in parsed_response:
                    for rule in parsed_response["rules"]:
                        if "coefficients" in rule:
                            is_valid, result = self.verify_rule(values, rule["coefficients"])
                            
                            verified_rule = {
                                "coefficients": rule["coefficients"],
                                "equation": rule.get("equation", ""),
                                "confidence": rule.get("confidence", 0.5),  # 保留LLM的信心度
                                "calculated_result": result,
                                "is_valid": is_valid  # 只使用我們計算的is_valid
                            }
                            
                            pattern_results["verified_rules"].append(verified_rule)
                            
                            if is_valid:
                                pattern_results["valid_count"] += 1
                                results["valid_rules"] += 1
                
                results["patterns"][pattern_name] = pattern_results
                
            except Exception as e:
                self.logger(f"分析模式 {pattern_name} 時發生錯誤: {e}")
                results["patterns"][pattern_name] = {
                    "values": values,
                    "error": str(e),
                    "verified_rules": [],
                    "valid_count": 0
                }
        
        # 計算合規比例
        total_rules = sum(len(p.get("verified_rules", [])) for p in results["patterns"].values())
        if total_rules > 0:
            results["rule_compliance_ratio"] = results["valid_rules"] / total_rules
        
        results["total_verified_rules"] = total_rules
        
        return results
    
    def save_results(self, results: Dict[str, Any], output_path: str):
        """保存結果到JSON文件"""
        try:
            # 確保目錄存在
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # 讀取現有結果（如果存在）
            existing_results = {}
            if os.path.exists(output_path):
                try:
                    with open(output_path, 'r', encoding='utf-8') as f:
                        existing_results = json.load(f)
                except:
                    existing_results = {}
            
            # 添加新結果
            program_key = results["program_path"]
            existing_results[program_key] = results
            
            # 保存結果
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(existing_results, f, ensure_ascii=False, indent=2)
            
            self.logger(f"結果已保存到: {output_path}")
            
        except Exception as e:
            self.logger(f"保存結果失敗: {e}")
            raise

def main():
    """主程式"""
    parser = argparse.ArgumentParser(description='數學規則偵測器')
    parser.add_argument('--file', required=True, help='Excel文件路徑')
    parser.add_argument('--sheet', help='工作表名稱')
    parser.add_argument('--output', default='tmp2/prompts/observed.json', help='輸出JSON文件路徑')
    parser.add_argument('--use-openai', action='store_true', help='使用OpenAI API')
    parser.add_argument('--window-height', type=int, default=5, help='窗格高度（分析幾個數值，預設5）')
    parser.add_argument('--window-width', type=int, default=1, help='窗格寬度（預設1）')
    
    args = parser.parse_args()
    
    # 初始化偵測器
    window_shape = (args.window_height, args.window_width)
    detector = MathRuleDetector(window_shape=window_shape)
    
    try:
        # 分析模式
        results = detector.analyze_patterns(
            file_path=args.file,
            sheet_name=args.sheet,
            use_openai=args.use_openai
        )
        
        # 保存結果
        detector.save_results(results, args.output)
        
        # 輸出摘要
        print(f"\n=== 分析結果摘要 ===")
        print(f"分析文件: {results['analyzed_file']}")
        print(f"工作表: {results['sheet_name']}")
        print(f"總模式數: {results['total_patterns']}")
        print(f"總驗證規則數: {results['total_verified_rules']}")
        print(f"有效規則數: {results['valid_rules']}")
        print(f"合規比例: {results['rule_compliance_ratio']:.2%}")
        print(f"結果已保存到: {args.output}")
        
    except Exception as e:
        print(f"執行失敗: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
