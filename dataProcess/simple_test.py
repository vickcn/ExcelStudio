#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
簡化測試腳本 - 不依賴語言模型API
直接使用已知的表格結構進行測試
"""

import os
import sys
import pandas as pd
from pathlib import Path

# 如果作為模組執行，使用相對導入
try:
    from .table_detector import TableDetector
except ImportError:
    # 如果直接執行，使用絕對導入
    sys.path.append(os.path.dirname(__file__))
    from table_detector import TableDetector

def test_basic_functionality():
    """測試基本功能（不使用語言模型）"""
    print("=== 基本功能測試 ===")
    
    detector = TableDetector()
    
    # 已知的正確結構
    expected_structure = {
        'correct_simple.xlsx': {
            'values_range': {'start_row': 4, 'end_row': 39, 'start_col': 2, 'end_col': 12},
            'index_range': {'start_row': 4, 'end_row': 39, 'start_col': 1, 'end_col': 1},  # 修正：索引從第4行開始
            'columns_range': {'start_row': 1, 'end_row': 1, 'start_col': 2, 'end_col': 12}
        }
    }
    
    for filename, expected in expected_structure.items():
        filepath = Path(__file__).parent.parent.parent / filename
        
        if not filepath.exists():
            print(f"檔案不存在: {filepath}")
            continue
        
        print(f"\n測試檔案: {filename}")
        
        # 讀取Excel檔案
        try:
            excel_data = pd.read_excel(filepath, sheet_name=None, header=None)
            df = list(excel_data.values())[0]  # 取第一個工作表
            
            print(f"  檔案形狀: {df.shape}")
            
            # 測試數字表格檢查
            is_numeric = detector._is_numeric_table(df)
            print(f"  是否為純數字表格: {is_numeric}")
            
            # 測試缺失值比例計算
            missing_ratio = detector._calculate_missing_ratio(df)
            print(f"  缺失值比例: {missing_ratio:.3f}")
            
            # 手動提取已知區域的DataFrame
            print("\n  手動提取表格區域:")
            
            # 提取數值區域 (B4:L39 -> 第4-39行，第2-12欄)
            values_df = df.iloc[3:39, 1:12]  # 0-based索引
            print(f"    數值區域形狀: {values_df.shape}")
            
            # 提取欄位標題 (B1:L1 -> 第1行，第2-12欄)
            columns_df = df.iloc[0:1, 1:12]
            print(f"    欄位標題: {list(columns_df.iloc[0])}")
            
            # 提取索引 (A4:A39 -> 第4-39行，第1欄) - 修正為與數值區域對應
            index_df = df.iloc[3:39, 0:1]  # 改為從第4行開始，與數值區域對應
            print(f"    索引前5個: {list(index_df.iloc[:5, 0])}")
            print(f"    索引區域形狀: {index_df.shape}")
            
            # 建立完整的DataFrame
            result_df = values_df.copy()
            result_df.columns = columns_df.iloc[0].values
            result_df.index = index_df.iloc[:, 0].values
            
            print(f"    最終DataFrame形狀: {result_df.shape}")
            print("    前3行預覽:")
            print(result_df.head(3).to_string(max_cols=6))
            
        except Exception as e:
            print(f"  處理檔案時發生錯誤: {e}")
            import traceback
            traceback.print_exc()

def test_sheet_to_text():
    """測試工作表轉文字功能"""
    print("\n=== 工作表轉文字測試 ===")
    
    detector = TableDetector()
    
    filepath = Path(__file__).parent.parent.parent / 'correct_simple.xlsx'
    
    if filepath.exists():
        try:
            excel_data = pd.read_excel(filepath, sheet_name=None, header=None)
            df = list(excel_data.values())[0]
            
            # 只轉換前5行5欄作為示例
            sample_df = df.iloc[:5, :5]
            text_content = detector._sheet_to_text(sample_df)
            
            print("前5行5欄的文字格式:")
            print(text_content[:500] + "..." if len(text_content) > 500 else text_content)
            
        except Exception as e:
            print(f"轉換失敗: {e}")

def main():
    """主函數"""
    print("開始簡化測試")
    
    try:
        test_basic_functionality()
        test_sheet_to_text()
        
        print("\n=== 簡化測試完成 ===")
        
    except Exception as e:
        print(f"測試過程中發生錯誤: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
