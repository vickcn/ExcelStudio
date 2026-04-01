#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用表格偵測器 - 不依賴預知範圍的自動偵測
"""

import os
import json
import pandas as pd
import numpy as np
import openai
import requests
from typing import Dict, List, Any, Optional, Tuple
from dotenv import load_dotenv
import argparse
from pathlib import Path

# 載入環境變數
load_dotenv()

class UniversalTableDetector:
    """通用表格偵測器 - 自動偵測未知範圍"""
    
    def __init__(self):
        """初始化偵測器"""
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
        
        # 建立 session
        self.session = requests.Session()
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
                    if content:
                        prompts[prompt_file.stem] = json.loads(content)
                    else:
                        print(f"跳過空的prompt檔案: {prompt_file.name}")
            except json.JSONDecodeError as e:
                print(f"無法載入prompt檔案 {prompt_file.name}: {e}")
            except Exception as e:
                print(f"載入prompt檔案 {prompt_file.name} 時發生錯誤: {e}")
        
        return prompts

    def _detect_fulltable_from_wb(
        self,
        excel_file: str,
        data_only: bool = False,
        sheet_name: Optional[str] = None,
        ret: Optional[Dict[str, Any]] = None,
    ) -> pd.DataFrame:
        """
        從 Excel 檔案中提取工作表（值為主，公式另行處理）。

        Return:
        - df: DataFrame（單一 sheet）
        """
        df = pd.read_excel(excel_file, header=None, sheet_name=sheet_name)
        # 若未指定 sheet，pandas 會回傳 dict；此處取第一張
        if isinstance(df, dict):
            first_name = next(iter(df.keys()), None)
            if ret is not None and isinstance(ret, dict):
                ret["sheet_name"] = first_name
            df = df.get(first_name) if first_name is not None else pd.DataFrame()
        return df

    def _detect_tables_from_full(
        self,
        df: pd.DataFrame,
        table_mode: str,
        use_llm: bool = False,
        prefer_local: bool = False,
    ) -> List[Dict[str, Any]]:
        # 分析數據分布
        data_regions = self._analyze_data_distribution(df, table_mode)
        
        # 對每個數據區域進行詳細分析
        results = []
        for i, region in enumerate(data_regions, 1):
            print(f"\n分析數據區域 {i}: {region}")
            
            if use_llm:
                # 使用LLM進行精確分析
                table_info = self._analyze_with_llm(df, region, table_mode, prefer_local)
            else:
                # 使用規則式分析
                table_info = self._analyze_table_structure(df, region, table_mode)
            
            if table_info:
                table_info['table_id'] = i
                results.append(table_info)

        return results

    def detect_tables_by_analysis(self, excel_file: str, table_mode: str = 'pure_numeric', use_llm: bool = False, prefer_local: bool = False, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """通過數據分析自動偵測表格"""
        print(f"開始分析表格結構: {excel_file}")
        print(f"偵測模式: {table_mode}")
        
        # 讀取Excel檔案
        try:
            # df = pd.read_excel(excel_file, header=None)
            df = self._detect_fulltable_from_wb(excel_file, sheet_name=sheet_name)
            print(f"檔案尺寸: {df.shape[0]}行 x {df.shape[1]}欄")
        except Exception as e:
            print(f" 無法讀取Excel檔案: {e}")
            return []
        
        results = self._detect_tables_from_full(df, table_mode, use_llm=use_llm, prefer_local=prefer_local)
        
        return results
    
    def _analyze_data_distribution(self, df: pd.DataFrame, table_mode: str) -> List[Dict[str, Any]]:
        """分析數據分布，找出可能的表格區域"""
        print(" 分析數據分布...")
        
        # 創建數據密度矩陣
        density_matrix = np.zeros(df.shape)
        
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                value = df.iloc[i, j]
                if self._is_valid_data(value, table_mode):
                    density_matrix[i, j] = 1
        
        # 尋找連續的高密度區域
        regions = self._find_dense_regions(density_matrix, min_size=(3, 3))
        
        print(f"找到 {len(regions)} 個潛在數據區域")
        return regions
    
    def _is_valid_data(self, value, table_mode: str) -> bool:
        """判斷是否為有效數據"""
        if pd.isna(value) or value == '' or value is None:
            return False
        
        if table_mode == 'pure_numeric':
            # 純數字模式
            if isinstance(value, (int, float)):
                return True
            elif isinstance(value, str):
                # 允許包含特殊符號的數字
                try:
                    # 移除常見的非數字符號
                    clean_value = str(value).replace('±', '').replace(',', '').strip()
                    float(clean_value)
                    return True
                except:
                    return False
        else:
            # 廣義模式 - 任何非空值都算有效
            return True
        
        return False
    
    def _find_dense_regions(self, density_matrix: np.ndarray, min_size: Tuple[int, int] = (3, 3)) -> List[Dict[str, Any]]:
        """尋找密集區域"""
        regions = []
        rows, cols = density_matrix.shape
        visited = np.zeros_like(density_matrix, dtype=bool)
        
        for i in range(rows):
            for j in range(cols):
                if density_matrix[i, j] == 1 and not visited[i, j]:
                    # 使用連通區域分析
                    region = self._flood_fill(density_matrix, visited, i, j)
                    
                    if region and self._is_valid_region(region, min_size):
                        regions.append(region)
        
        return regions
    
    def _flood_fill(self, density_matrix: np.ndarray, visited: np.ndarray, start_row: int, start_col: int) -> Optional[Dict[str, Any]]:
        """洪水填充算法找連通區域"""
        rows, cols = density_matrix.shape
        stack = [(start_row, start_col)]
        region_cells = []
        
        while stack:
            r, c = stack.pop()
            
            if (r < 0 or r >= rows or c < 0 or c >= cols or 
                visited[r, c] or density_matrix[r, c] == 0):
                continue
            
            visited[r, c] = True
            region_cells.append((r, c))
            
            # 檢查8個方向（包括對角線）
            for dr in [-1, 0, 1]:
                for dc in [-1, 0, 1]:
                    if dr == 0 and dc == 0:
                        continue
                    stack.append((r + dr, c + dc))
        
        if not region_cells:
            return None
        
        # 計算邊界框
        rows_list = [r for r, c in region_cells]
        cols_list = [c for r, c in region_cells]
        
        return {
            'start_row': min(rows_list) + 1,  # 轉換為1-based索引
            'end_row': max(rows_list) + 1,
            'start_col': min(cols_list) + 1,
            'end_col': max(cols_list) + 1,
            'cell_count': len(region_cells)
        }
    
    def _is_valid_region(self, region: Dict[str, Any], min_size: Tuple[int, int]) -> bool:
        """判斷區域是否有效"""
        height = region['end_row'] - region['start_row'] + 1
        width = region['end_col'] - region['start_col'] + 1
        
        return height >= min_size[0] and width >= min_size[1]
    
    def _analyze_table_structure(self, df: pd.DataFrame, region: Dict[str, Any], table_mode: str) -> Optional[Dict[str, Any]]:
        """分析表格結構，尋找標題和索引"""
        print(f"   分析表格結構...")
        
        # 數據區域（0-based索引）
        data_start_row = region['start_row'] - 1
        data_end_row = region['end_row'] - 1
        data_start_col = region['start_col'] - 1
        data_end_col = region['end_col'] - 1
        
        # 尋找列標題（在數據區域上方）
        columns_range = self._find_column_headers(df, region)
        
        # 尋找行索引（在數據區域左側）
        index_range = self._find_row_index(df, region)
        
        # 生成描述
        description = f"數據表格 ({region['end_row'] - region['start_row'] + 1}x{region['end_col'] - region['start_col'] + 1})"
        
        return {
            'description': description,
            'sheet_name': '工作表1',  # 簡化處理
            'values_range': {
                'start_row': region['start_row'],
                'end_row': region['end_row'],
                'start_col': region['start_col'],
                'end_col': region['end_col']
            },
            'columns_range': columns_range,
            'index_range': index_range
        }
    
    def _find_column_headers(self, df: pd.DataFrame, region: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """尋找列標題"""
        data_start_row = region['start_row'] - 1
        data_start_col = region['start_col'] - 1
        data_end_col = region['end_col'] - 1
        
        # 檢查數據區域上方的幾行
        for header_row in range(max(0, data_start_row - 3), data_start_row):
            # 檢查這一行是否可能是標題
            header_values = df.iloc[header_row, data_start_col:data_end_col + 1]
            
            # 如果大部分都不是空值，可能是標題
            non_empty_count = sum(1 for v in header_values if not pd.isna(v) and v != '' and v is not None)
            if non_empty_count >= len(header_values) * 0.5:  # 至少50%非空
                return {
                    'start_row': header_row + 1,
                    'end_row': header_row + 1,
                    'start_col': data_start_col + 1,
                    'end_col': data_end_col + 1
                }
        
        return None
    
    def _find_row_index(self, df: pd.DataFrame, region: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """尋找行索引"""
        data_start_row = region['start_row'] - 1
        data_end_row = region['end_row'] - 1
        data_start_col = region['start_col'] - 1
        
        # 檢查數據區域左側的幾列
        for index_col in range(max(0, data_start_col - 3), data_start_col):
            # 檢查這一列是否可能是索引
            index_values = df.iloc[data_start_row:data_end_row + 1, index_col]
            
            # 如果大部分都不是空值，可能是索引
            non_empty_count = sum(1 for v in index_values if not pd.isna(v) and v != '' and v is not None)
            if non_empty_count >= len(index_values) * 0.7:  # 至少70%非空
                return {
                    'start_row': data_start_row + 1,
                    'end_row': data_end_row + 1,
                    'start_col': index_col + 1,
                    'end_col': index_col + 1
                }
        
        return None
    
    def _analyze_with_llm(self, df: pd.DataFrame, region: Dict[str, Any], table_mode: str, prefer_local: bool = False) -> Optional[Dict[str, Any]]:
        """使用LLM分析表格結構"""
        print(f"   使用LLM分析表格結構...")
        
        # 提取區域周圍的數據供LLM分析
        context_data = self._extract_context_data(df, region)
        
        # 準備prompt
        if table_mode == 'pure_numeric':
            prompt_template = self.prompts.get('pure_numeric_table')
        else:
            prompt_template = self.prompts.get('general_table')
        
        if not prompt_template:
            print(f" 找不到prompt模板: {table_mode}")
            return self._analyze_table_structure(df, region, table_mode)
        
        # 構建訊息
        messages = [
            {"role": "system", "content": prompt_template['system_prompt']},
            {"role": "user", "content": prompt_template['user_prompt'].format(sheet_content=context_data)}
        ]
        
        # 呼叫LLM
        response = self._call_llm(messages, prefer_local)
        if not response:
            print(f" LLM分析失敗，使用規則式分析")
            return self._analyze_table_structure(df, region, table_mode)
        
        # 解析LLM回應
        try:
            llm_results = json.loads(response)
            if isinstance(llm_results, list) and len(llm_results) > 0:
                result = llm_results[0]  # 取第一個結果
                result['description'] = result.get('description', '由LLM偵測的表格')
                result['sheet_name'] = '工作表1'
                return result
        except json.JSONDecodeError as e:
            print(f" LLM回應JSON解析失敗: {e}")
        
        # LLM失敗時回退到規則式分析
        return self._analyze_table_structure(df, region, table_mode)
    
    def _extract_context_data(self, df: pd.DataFrame, region: Dict[str, Any]) -> str:
        """提取區域周圍的數據供LLM分析"""
        # 擴展區域以包含上下文
        start_row = max(0, region['start_row'] - 5)
        end_row = min(df.shape[0], region['end_row'] + 3)
        start_col = max(0, region['start_col'] - 3)
        end_col = min(df.shape[1], region['end_col'] + 2)
        
        context_df = df.iloc[start_row:end_row, start_col:end_col]
        
        # 轉換為文字格式
        text_lines = []
        for i in range(context_df.shape[0]):
            row_data = []
            for j in range(context_df.shape[1]):
                cell_value = context_df.iloc[i, j]
                if pd.isna(cell_value) or cell_value == '':
                    cell_value = ''
                
                # 計算實際的Excel座標
                excel_row = start_row + i + 1
                excel_col = chr(ord('A') + start_col + j) if start_col + j < 26 else f"A{chr(ord('A') + start_col + j - 26)}"
                row_data.append(f"{excel_col}{excel_row}: {cell_value}")
            
            text_lines.append(", ".join(row_data))
        
        return "\n".join(text_lines)
    
    def _call_llm(self, messages: List[Dict[str, str]], prefer_local: bool = False) -> Optional[str]:
        """呼叫LLM"""
        if prefer_local:
            print(" 優先使用本地LLM")
            # 優先使用本地LLM
            result = self._call_local_llm(messages)
            if result:
                print(" 本地LLM回應成功")
                return result
            
            print(" 本地LLM失敗，嘗試OpenAI")
            # 本地LLM失敗時使用OpenAI
            if self.openai_client:
                result = self._call_openai(messages)
                if result:
                    print(" OpenAI回應成功")
                    return result
        else:
            print(" 優先使用OpenAI")
            # 優先使用OpenAI
            if self.openai_client:
                result = self._call_openai(messages)
                if result:
                    print(" OpenAI回應成功")
                    return result
                print(" OpenAI失敗，嘗試本地LLM")
            else:
                print(" OpenAI客戶端未初始化，使用本地LLM")
            
            # OpenAI失敗時使用本地LLM
            result = self._call_local_llm(messages)
            if result:
                print(" 本地LLM回應成功")
                return result
        
        return None
    
    def _call_local_llm(self, messages: List[Dict[str, str]]) -> Optional[str]:
        """呼叫本地LLM"""
        try:
            payload = {
                'model': self.llm_model_name,
                'messages': messages,
                'temperature': 0.1,
                'max_tokens': 4000
            }
            
            headers = {}
            if self.api_key:
                headers['Authorization'] = f'Bearer {self.api_key}'
            
            response = self.session.post(
                self.llm_api_url,
                json=payload,
                headers=headers,
                timeout=60
            )
            
            response.raise_for_status()
            result = response.json()
            
            if 'choices' in result and len(result['choices']) > 0:
                return result['choices'][0]['message']['content']
            elif 'response' in result:
                return result['response']
            
        except Exception as e:
            print(f" 本地LLM呼叫失敗: {e}")
        
        return None
    
    def _call_openai(self, messages: List[Dict[str, str]]) -> Optional[str]:
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
            print(f" OpenAI API呼叫失敗: {e}")
        
        return None
    
    def _extract_dataframe_from_full(self, df: pd.DataFrame, result: Dict[str, Any]) -> pd.DataFrame:
        """從完整表格擷取單一 DataFrame 區塊。"""
        val_range = result['values_range']
        values_df = df.iloc[
            val_range['start_row']-1:val_range['end_row'],
            val_range['start_col']-1:val_range['end_col']
        ].copy()

        # 設定欄位名稱
        if result.get('columns_range'):
            col_range = result['columns_range']
            columns_df = df.iloc[
                col_range['start_row']-1:col_range['end_row'],
                col_range['start_col']-1:col_range['end_col']
            ]
            if len(columns_df) > 0:
                values_df.columns = columns_df.iloc[0].values

        # 設定索引
        if result.get('index_range'):
            idx_range = result['index_range']
            index_df = df.iloc[
                idx_range['start_row']-1:idx_range['end_row'],
                idx_range['start_col']-1:idx_range['end_col']
            ]
            if len(index_df) == len(values_df):
                values_df.index = index_df.iloc[:, 0].values
            else:
                print(f"  索引長度不匹配: values={len(values_df)}, index={len(index_df)}")
                print(f"    使用預設索引")

        return values_df

    def extract_dataframes(self, excel_file: str, results: List[Dict[str, Any]]) -> List[pd.DataFrame]:
        """提取DataFrame（從檔案）"""
        df = pd.read_excel(excel_file, header=None)
        return self.extract_dataframes_from_full(df, results)

    def extract_dataframes_from_full(self, df: pd.DataFrame, results: List[Dict[str, Any]]) -> List[pd.DataFrame]:
        """提取DataFrame（從完整表格）"""
        dataframes = []
        for result in results:
            dataframes.append(self._extract_dataframe_from_full(df, result))
        return dataframes

def main():
    """主函數"""
    parser = argparse.ArgumentParser(description='通用表格偵測器')
    parser.add_argument('excel_file', help='Excel檔案路徑')
    parser.add_argument('--mode', choices=['pure_numeric', 'general'], 
                       default='pure_numeric', help='表格偵測模式')
    parser.add_argument('--use-llm', action='store_true', 
                       help='使用LLM進行精確分析（預設使用規則式分析）')
    parser.add_argument('--prefer-local', action='store_true',
                       help='優先使用本地LLM（與--use-llm一起使用時有效）')
    
    args = parser.parse_args()
    
    # 執行偵測
    detector = UniversalTableDetector()
    results = detector.detect_tables_by_analysis(args.excel_file, args.mode, args.use_llm, args.prefer_local)
    
    print(f"\n偵測完成，共找到 {len(results)} 個表格")
    
    for i, result in enumerate(results, 1):
        print(f"\n表格 {i}:")
        print(f"  工作表: {result.get('sheet_name', 'N/A')}")
        print(f"  描述: {result['description']}")
        print(f"  數值區域: {result['values_range']}")
        print(f"  欄位區域: {result.get('columns_range', 'None')}")
        print(f"  索引區域: {result.get('index_range', 'None')}")
    
    # 提取DataFrame示例
    if results:
        print(f"\n提取DataFrame示例:")
        try:
            dataframes = detector.extract_dataframes(args.excel_file, results)
            for i, df in enumerate(dataframes, 1):
                print(f"\n表格 {i} DataFrame:")
                print(f"形狀: {df.shape}")
                print(df.head())

                print(f"columns: {','.join(df.columns)}")
                print(f"index: {','.join(df.index)}")
        except Exception as e:
            print(f" DataFrame提取失敗: {e}")

if __name__ == "__main__":
    main()
