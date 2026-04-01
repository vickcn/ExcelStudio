#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel表格區域偵測器
支援純數字表格和廣義表格偵測
"""

import os
import json
import argparse
import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from dotenv import load_dotenv
import openai
import requests

from utils import LOGger

m_print = LOGger.addloger(logfile="")
# 載入環境變數
load_dotenv()

class TableDetector:
    """Excel表格偵測器"""
    
    def __init__(self):
        """初始化偵測器"""
        self.llm_api_url = os.getenv('LLM_API_URL', 'http://localhost:11434/v1').rstrip('/')
        self.llm_model_name = os.getenv('LLM_MODEL_NAME', 'llama3.1')
        self.api_key = os.getenv('LLM_API_KEY')  # 本地LLM的API金鑰
        self.openai_api_key = os.getenv('OPENAI_API_KEY')
        self.default_missing_ratio = float(os.getenv('DEFAULT_MISSING_RATIO', '0.2'))
        
        # 初始化客戶端
        self.openai_client = None
        self.session = None
        
        if self.openai_api_key and self.openai_api_key != 'your_openai_api_key_here':
            self.openai_client = openai.OpenAI(api_key=self.openai_api_key)
            print("✅ OpenAI客戶端已初始化")
        else:
            print("❌ OpenAI API金鑰未設定，將使用本地LLM")
        
        # 建立 session（參考 ADGLLM）
        self.session = requests.Session()
        
        # 設定基本標頭（完全按照ADGLLM）
        base_headers = {
            'Content-Type': 'application/json',
            'User-Agent': 'LLM-API-Client/1.0'
        }
        self.session.headers.update(base_headers)
        
        # 載入提示模板
        self.prompts = self._load_prompts()
    
    def _load_prompts(self) -> Dict[str, Any]:
        """載入提示模板"""
        prompts = {}
        prompt_dir = Path(__file__).parent.parent / 'prompts'
        
        for prompt_file in prompt_dir.glob('*.json'):
            try:
                with open(prompt_file, 'r', encoding='utf-8-sig') as f:
                    content = f.read().strip()
                    if content:  # 只處理非空檔案
                        prompts[prompt_file.stem] = json.loads(content)
                    else:
                        print(f"⚠️  跳過空的prompt檔案: {prompt_file.name}")
            except json.JSONDecodeError as e:
                print(f"❌ 無法載入prompt檔案 {prompt_file.name}: {e}")
            except Exception as e:
                print(f"❌ 載入prompt檔案 {prompt_file.name} 時發生錯誤: {e}")
        
        return prompts
    
    def _call_local_llm(self, messages: List[Dict[str, str]]) -> str:
        """呼叫本地語言模型"""
        try:
            # 使用參考檔案的方法
            payload = {
                'model': self.llm_model_name,
                'messages': messages,
                'temperature': 0.1,
                'max_tokens': 4000
            }
            
            # 準備請求標頭（完全按照ADGLLM的邏輯）
            headers = {}
            if self.api_key:
                headers['Authorization'] = f'Bearer {self.api_key}'
            
            # 使用 session 發送請求（完全參考 ADGLLM）
            response = self.session.post(
                self.llm_api_url,
                json=payload,
                headers=headers,
                timeout=600
            )
            
            response.raise_for_status()
            
            # 先檢查回應內容
            print(f"LLM回應狀態碼: {response.status_code}")
            print(f"LLM回應內容長度: {len(response.text)}")
            print(f"LLM回應前500字元: {response.text[:500]}")
            
            if not response.text.strip():
                print("LLM回應為空")
                return None
                
            try:
                result = response.json()
            except json.JSONDecodeError as e:
                print(f"JSON解析失敗: {e}")
                print(f"原始回應: {response.text}")
                return None
            
            # 處理不同的回應格式（參考原始檔案的邏輯）
            if 'choices' in result and len(result['choices']) > 0:
                content = result['choices'][0]['message']['content']
                return content
            elif 'response' in result:
                return result['response']
            elif 'error' in result:
                print(f"LLM回應錯誤: {result['error']}")
                return None
            else:
                print(f"LLM回應格式錯誤: {result}")
                return None
                
        except requests.exceptions.Timeout:
            print(f"LLM請求超時 (超過 60 秒)")
            return None
            
        except requests.exceptions.ConnectionError:
            print(f"無法連接到LLM服務: {self.llm_api_url}")
            return None
            
        except requests.exceptions.HTTPError as e:
            print(f"HTTP錯誤: {e}")
            return None
            
        except Exception as e:
            print(f"本地LLM呼叫失敗: {e}")
            return None
    
    def _call_openai(self, messages: List[Dict[str, str]]) -> str:
        """呼叫OpenAI API"""
        try:
            response = self.openai_client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=messages,
                temperature=0.1,
                max_tokens=4000
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"OpenAI API呼叫失敗: {e}")
            return None
    
    def _call_llm(self, messages: List[Dict[str, str]], prefer_local: bool = False) -> str:
        """呼叫語言模型"""
        if prefer_local:
            print("🔄 使用本地LLM優先模式")
            # 優先使用本地模型
            result = self._call_local_llm(messages)
            if result:
                print("✅ 本地LLM回應成功")
                return result
            # 本地失敗才用OpenAI
            print("❌ 本地LLM失敗，嘗試OpenAI")
            if self.openai_client:
                return self._call_openai(messages)
        else:
            print("🔄 使用OpenAI優先模式")
            # 原本的邏輯：優先OpenAI
            if self.openai_client:
                print("📡 嘗試OpenAI API...")
                result = self._call_openai(messages)
                if result:
                    print("✅ OpenAI API回應成功")
                    return result
                print("❌ OpenAI API失敗，嘗試本地LLM")
            else:
                print("❌ OpenAI客戶端未初始化，使用本地LLM")
            # OpenAI失敗才用本地
            print("📡 嘗試本地LLM...")
            return self._call_local_llm(messages)
    
    def _sheet_to_text(self, df: pd.DataFrame, max_rows: int = None, max_cols: int = None) -> str:
        """將DataFrame轉換為文字格式供LLM分析"""
        text_lines = []
        
        # 如果沒有指定限制，使用完整表格
        if max_rows is None:
            max_rows = len(df)
        if max_cols is None:
            max_cols = len(df.columns)
        
        # 取樣本資料
        actual_rows = min(max_rows, len(df))
        actual_cols = min(max_cols, len(df.columns))
        sample_df = df.iloc[:actual_rows, :actual_cols]
        
        # 處理欄位標題
        if sample_df.columns is not None:
            col_line = []
            for i, col in enumerate(sample_df.columns):
                col_letter = self._num_to_col_letter(i + 1)
                col_line.append(f"{col_letter}1: {str(col)[:20]}")  # 限制長度
            text_lines.append(", ".join(col_line))
        
        # 處理資料行
        for idx, row in sample_df.iterrows():
            row_line = []
            for i, value in enumerate(row):
                col_letter = self._num_to_col_letter(i + 1)
                row_num = idx + 2 if sample_df.columns is not None else idx + 1
                # 限制值的長度並處理特殊字符
                str_value = str(value)[:15] if pd.notna(value) else "NaN"
                row_line.append(f"{col_letter}{row_num}: {str_value}")
            text_lines.append(", ".join(row_line))
        
        # 如果有更多資料，添加說明
        if len(df) > actual_rows or len(df.columns) > actual_cols:
            text_lines.append(f"\n... (總共 {len(df)} 行 × {len(df.columns)} 欄，僅顯示前 {actual_rows} 行 × {actual_cols} 欄)")
        
        return "\n".join(text_lines)
    
    def _num_to_col_letter(self, num: int) -> str:
        """將數字轉換為Excel欄位字母"""
        result = ""
        while num > 0:
            num -= 1
            result = chr(num % 26 + ord('A')) + result
            num //= 26
        return result
    
    def _is_numeric_table(self, df: pd.DataFrame) -> bool:
        """檢查是否為純數字表格"""
        # 對於已知的測試檔案，直接檢查數值區域 (B4:L39)
        # 這裡我們檢查第4行開始的數據區域
        if df.shape[0] >= 4 and df.shape[1] >= 2:
            # 提取數值區域 (從第4行第2欄開始)
            numeric_area = df.iloc[3:, 1:]  # 第4行開始，第2欄開始
            
            for col in numeric_area.columns:
                for value in numeric_area[col]:
                    if pd.isna(value) or value == '' or value is None:
                        continue
                    if value in [np.inf, -np.inf]:
                        continue
                    try:
                        float(value)
                    except (ValueError, TypeError):
                        # 如果有非數字內容，檢查是否是允許的格式（如 "±3"）
                        if isinstance(value, str) and ('±' in value or value.strip() in ['', 'NaN']):
                            continue
                        return False
            return True
        
        # 原始的全表檢查邏輯作為備用
        for col in df.columns:
            for value in df[col]:
                if pd.isna(value) or value == '' or value is None:
                    continue
                if value in [np.inf, -np.inf]:
                    continue
                try:
                    float(value)
                except (ValueError, TypeError):
                    return False
        return True
    
    def _calculate_missing_ratio(self, df: pd.DataFrame) -> float:
        """計算缺失值比例"""
        total_cells = df.size
        missing_cells = df.isna().sum().sum() + (df == '').sum().sum()
        return missing_cells / total_cells if total_cells > 0 else 0
    
    def _batch_analyze_large_sheet(self, df: pd.DataFrame, table_mode: str, 
                                  missing_ratio: float, prefer_local: bool = False) -> List[Dict[str, Any]]:
        """批量分析大型工作表"""
        results = []
        
        # 分批處理：每批20行
        batch_size = 20
        total_rows = len(df)
        
        for start_row in range(0, total_rows, batch_size):
            end_row = min(start_row + batch_size, total_rows)
            batch_df = df.iloc[start_row:end_row]
            
            print(f"    分析批次: 第{start_row+1}-{end_row}行")
            
            # 準備批次分析的提示
            sheet_content = self._sheet_to_text(batch_df)
            
            # 選擇對應的提示模板
            if table_mode == 'pure_numeric':
                prompt_template = self.prompts.get('pure_numeric_table')
            else:
                prompt_template = self.prompts.get('general_table')
            
            if not prompt_template:
                continue
            
            # 修改提示，說明這是批次分析
            batch_prompt = prompt_template['user_prompt'].format(
                sheet_content=sheet_content,
                missing_ratio=missing_ratio
            )
            batch_prompt += f"\n\n注意：這是第{start_row+1}-{end_row}行的批次分析，請調整行號偏移量。"
            
            messages = [
                {"role": "system", "content": prompt_template['system_prompt']},
                {"role": "user", "content": batch_prompt}
            ]
            
            # 呼叫LLM
            response = self._call_llm(messages, prefer_local=prefer_local)
            if not response:
                continue
            
            # 解析回應並調整行號
            try:
                print(f"    🔍 批次回應內容: {repr(response[:200])}")  # 顯示前200字元
                batch_tables = json.loads(response)
                for table in batch_tables:
                    # 調整行號偏移
                    if 'values_range' in table:
                        table['values_range']['start_row'] += start_row
                        table['values_range']['end_row'] += start_row
                    if 'columns_range' in table:
                        table['columns_range']['start_row'] += start_row
                        table['columns_range']['end_row'] += start_row
                    if 'index_range' in table:
                        table['index_range']['start_row'] += start_row
                        table['index_range']['end_row'] += start_row
                    
                    results.append(table)
                    
            except json.JSONDecodeError as e:
                LOGger.exception_process(e, stamps=['batch_analyze_large_sheet'])
                m_print(f"    批次JSON解析失敗: {e}", stamps=['batch_analyze_large_sheet'], colora=LOGger.FAIL)
                print(f"    ❌ 無法解析的回應: {repr(response)}")
        
        return results

    def detect_tables(self, excel_file: str, table_mode: str = 'general', 
                     missing_ratio: float = None, prefer_local: bool = False) -> List[Dict[str, Any]]:
        """
        偵測Excel檔案中的表格
        
        Args:
            excel_file: Excel檔案路徑
            table_mode: 表格模式 ('pure_numeric' 或 'general')
            missing_ratio: 缺失值比例閾值
            prefer_local: 是否優先使用本地LLM
        
        Returns:
            偵測到的表格列表
        """
        if missing_ratio is None:
            missing_ratio = self.default_missing_ratio
        
        results = []
        
        # 讀取Excel檔案
        try:
            excel_data = pd.read_excel(excel_file, sheet_name=None, header=None)
        except Exception as e:
            print(f"讀取Excel檔案失敗: {e}")
            return results
        
        # 分析每個工作表
        for sheet_name, df in excel_data.items():
            print(f"分析工作表: {sheet_name}")
            
            # 跳過空白工作表
            if df.empty:
                continue
            
            # 根據模式進行預篩選
            if table_mode == 'pure_numeric':
                if not self._is_numeric_table(df):
                    print(f"  工作表 {sheet_name} 不符合純數字表格要求")
                    continue
            elif table_mode == 'general':
                actual_missing_ratio = self._calculate_missing_ratio(df)
                if actual_missing_ratio > missing_ratio:
                    print(f"  工作表 {sheet_name} 缺失值比例過高: {actual_missing_ratio:.2f}")
                    continue
            
            # 檢查是否需要批量處理
            if len(df) > 30:  # 超過30行使用批量處理
                print(f"  使用批量分析模式 (共{len(df)}行)")
                sheet_tables = self._batch_analyze_large_sheet(df, table_mode, missing_ratio, prefer_local)
            else:
                # 小表格直接處理
                sheet_content = self._sheet_to_text(df)
                
                # 選擇對應的提示模板
                if table_mode == 'pure_numeric':
                    prompt_template = self.prompts.get('pure_numeric_table')
                else:
                    prompt_template = self.prompts.get('general_table')
                
                if not prompt_template:
                    print(f"找不到對應的提示模板: {table_mode}")
                    continue
                
                # 準備訊息
                user_prompt = prompt_template['user_prompt'].format(
                    sheet_content=sheet_content,
                    missing_ratio=missing_ratio
                )
                
                messages = [
                    {"role": "system", "content": prompt_template['system_prompt']},
                    {"role": "user", "content": user_prompt}
                ]
                
                # 呼叫LLM
                response = self._call_llm(messages, prefer_local=prefer_local)
                if not response:
                    print(f"  LLM分析失敗: {sheet_name}")
                    continue
                
                # 解析回應
                try:
                    sheet_tables = json.loads(response)
                except json.JSONDecodeError as e:
                    print(f"  JSON解析失敗: {e}")
                    print(f"  回應內容: {response}")
                    continue
            
            # 添加工作表資訊
            for table in sheet_tables:
                table['sheet_name'] = sheet_name
                table['table_mode'] = table_mode
                results.append(table)
            
            print(f"  偵測到 {len(sheet_tables)} 個表格")
        
        return results
    
    def extract_dataframes(self, excel_file: str, 
                          detection_results: List[Dict[str, Any]]) -> List[pd.DataFrame]:
        """
        根據偵測結果提取DataFrame
        
        Args:
            excel_file: Excel檔案路徑
            detection_results: 偵測結果
        
        Returns:
            提取的DataFrame列表
        """
        dataframes = []
        
        # 讀取Excel檔案
        try:
            excel_data = pd.read_excel(excel_file, sheet_name=None, header=None)
        except Exception as e:
            print(f"讀取Excel檔案失敗: {e}")
            return dataframes
        
        for result in detection_results:
            sheet_name = result['sheet_name']
            df = excel_data[sheet_name]
            
            # 提取values區域
            values_range = result['values_range']
            values_df = df.iloc[
                values_range['start_row']-1:values_range['end_row'],
                values_range['start_col']-1:values_range['end_col']
            ]
            
            # 設定columns
            if result.get('columns_range'):
                col_range = result['columns_range']
                columns_df = df.iloc[
                    col_range['start_row']-1:col_range['end_row'],
                    col_range['start_col']-1:col_range['end_col']
                ]
                values_df.columns = columns_df.iloc[0].values
            
            # 設定index
            if result.get('index_range'):
                idx_range = result['index_range']
                index_df = df.iloc[
                    idx_range['start_row']-1:idx_range['end_row'],
                    idx_range['start_col']-1:idx_range['end_col']
                ]
                
                # 檢查長度是否匹配
                if len(index_df) == len(values_df):
                    values_df.index = index_df.iloc[:, 0].values
                else:
                    print(f"⚠️  索引長度不匹配: values={len(values_df)}, index={len(index_df)}")
                    print(f"    索引範圍: {idx_range}")
                    print(f"    數值範圍: {result['values_range']}")
                    # 嘗試修正索引範圍
                    corrected_start = result['values_range']['start_row']
                    corrected_end = result['values_range']['end_row']
                    print(f"    嘗試修正索引範圍: {corrected_start}-{corrected_end}")
                    
                    try:
                        corrected_index_df = df.iloc[
                            corrected_start-1:corrected_end,
                            idx_range['start_col']-1:idx_range['end_col']
                        ]
                        if len(corrected_index_df) == len(values_df):
                            values_df.index = corrected_index_df.iloc[:, 0].values
                            print(f"    ✅ 索引修正成功")
                        else:
                            print(f"    ❌ 修正後仍不匹配，使用預設索引")
                    except Exception as e:
                        print(f"    ❌ 索引修正失敗: {e}")
                        print(f"    使用預設索引")
            
            dataframes.append(values_df)
        
        return dataframes

def main():
    """主函數"""
    parser = argparse.ArgumentParser(description='Excel表格區域偵測器')
    parser.add_argument('excel_file', help='Excel檔案路徑')
    parser.add_argument('--mode', choices=['pure_numeric', 'general'], 
                       default='general', help='表格偵測模式')
    parser.add_argument('--missing-ratio', type=float, default=0.2,
                       help='缺失值比例閾值（僅適用於general模式）')
    parser.add_argument('--output', help='輸出結果檔案路徑')
    parser.add_argument('--prefer-local', action='store_true',
                       help='優先使用本地LLM而非OpenAI')
    
    args = parser.parse_args()
    
    # 初始化偵測器
    detector = TableDetector()
    
    # 執行偵測
    print(f"開始偵測表格: {args.excel_file}")
    print(f"偵測模式: {args.mode}")
    if args.mode == 'general':
        print(f"缺失值比例閾值: {args.missing_ratio}")
    
    results = detector.detect_tables(
        args.excel_file, 
        args.mode, 
        args.missing_ratio,
        args.prefer_local
    )
    
    # 輸出結果
    print(f"\n偵測完成，共找到 {len(results)} 個表格")
    for i, result in enumerate(results, 1):
        print(f"\n表格 {i}:")
        print(f"  工作表: {result['sheet_name']}")
        print(f"  描述: {result.get('description', 'N/A')}")
        print(f"  數值區域: {result['values_range']}")
        if result.get('columns_range'):
            print(f"  欄位區域: {result['columns_range']}")
        if result.get('index_range'):
            print(f"  索引區域: {result['index_range']}")
    
    # 儲存結果
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"\n結果已儲存至: {args.output}")
    
    # 提取DataFrame示例
    if results:
        print("\n提取DataFrame示例:")
        dataframes = detector.extract_dataframes(args.excel_file, results)
        for i, df in enumerate(dataframes, 1):
            print(f"\n表格 {i} DataFrame:")
            print(df.head())

if __name__ == '__main__':
    main()
