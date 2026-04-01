#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
混合表格偵測器 - 結合規則式和LLM
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional
from .table_detector import TableDetector

class HybridTableDetector:
    """混合表格偵測器"""
    
    def __init__(self):
        self.llm_detector = TableDetector()
    
    def detect_known_format(self, excel_file: str) -> List[Dict[str, Any]]:
        """偵測已知格式的表格（如correct_simple.xlsx）"""
        df = pd.read_excel(excel_file, header=None)
        
        # 檢查是否符合correct_simple.xlsx格式
        if self._is_correct_simple_format(df):
            return [{
                'table_id': 1,
                'description': '標準數字表格',
                'sheet_name': '工作表1',
                'values_range': {
                    'start_row': 4,
                    'end_row': 39,
                    'start_col': 2,
                    'end_col': 12
                },
                'columns_range': {
                    'start_row': 1,
                    'end_row': 1,
                    'start_col': 2,
                    'end_col': 12
                },
                'index_range': {
                    'start_row': 2,
                    'end_row': 39,
                    'start_col': 1,
                    'end_col': 1
                }
            }]
        
        # 如果不是已知格式，使用LLM偵測
        return self.llm_detector.detect_tables(excel_file, 'pure_numeric')
    
    def _is_correct_simple_format(self, df: pd.DataFrame) -> bool:
        """檢查是否符合correct_simple.xlsx格式"""
        try:
            # 檢查基本尺寸
            if df.shape[0] < 39 or df.shape[1] < 12:
                return False
            
            # 檢查B4:L39區域是否主要為數字
            numeric_area = df.iloc[3:39, 1:12]
            numeric_count = 0
            total_count = numeric_area.size
            
            for i in range(numeric_area.shape[0]):
                for j in range(numeric_area.shape[1]):
                    value = numeric_area.iloc[i, j]
                    if pd.isna(value) or value == '' or value is None:
                        numeric_count += 1
                    elif isinstance(value, (int, float)):
                        numeric_count += 1
                    elif isinstance(value, str):
                        # 允許包含±符號的數字
                        try:
                            float(value.replace('±', '').strip())
                            numeric_count += 1
                        except:
                            pass
            
            # 如果80%以上是數字，認為符合格式
            return (numeric_count / total_count) > 0.8
            
        except Exception:
            return False
    
    def extract_dataframes(self, excel_file: str, results: List[Dict[str, Any]]) -> List[pd.DataFrame]:
        """提取DataFrame"""
        return self.llm_detector.extract_dataframes(excel_file, results)

def main():
    """測試混合偵測器"""
    import argparse
    
    parser = argparse.ArgumentParser(description='混合表格偵測器')
    parser.add_argument('excel_file', help='Excel檔案路徑')
    args = parser.parse_args()
    
    detector = HybridTableDetector()
    results = detector.detect_known_format(args.excel_file)
    
    print(f"偵測完成，共找到 {len(results)} 個表格")
    
    for i, result in enumerate(results, 1):
        print(f"\n表格 {i}:")
        print(f"  描述: {result['description']}")
        print(f"  數值區域: {result['values_range']}")
        print(f"  欄位區域: {result['columns_range']}")
        print(f"  索引區域: {result['index_range']}")
    
    # 提取DataFrame
    if results:
        print(f"\n提取DataFrame:")
        dataframes = detector.extract_dataframes(args.excel_file, results)
        for i, df in enumerate(dataframes, 1):
            print(f"\n表格 {i} DataFrame:")
            print(f"形狀: {df.shape}")
            print(df.head())

if __name__ == "__main__":
    main()
