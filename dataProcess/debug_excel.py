#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel檔案診斷腳本
檢查檔案內容和數據類型
"""

import pandas as pd
import numpy as np
from pathlib import Path

def debug_excel_file(filepath):
    """診斷Excel檔案內容"""
    print(f"=== 診斷檔案: {filepath} ===")
    
    try:
        # 讀取Excel檔案
        excel_data = pd.read_excel(filepath, sheet_name=None, header=None)
        
        for sheet_name, df in excel_data.items():
            print(f"\n工作表: {sheet_name}")
            print(f"形狀: {df.shape}")
            
            # 顯示前幾行
            print("\n前5行內容:")
            print(df.head())
            
            # 檢查數據類型
            print("\n數據類型:")
            print(df.dtypes)
            
            # 檢查非數字內容
            print("\n非數字內容檢查:")
            non_numeric_found = False
            
            for col_idx, col in enumerate(df.columns):
                for row_idx, value in enumerate(df[col]):
                    if pd.isna(value) or value == '' or value is None:
                        continue
                    if value in [np.inf, -np.inf]:
                        continue
                    
                    try:
                        float(value)
                    except (ValueError, TypeError):
                        print(f"  位置 ({row_idx+1}, {col_idx+1}): {repr(value)} (類型: {type(value)})")
                        non_numeric_found = True
                        
                        # 只顯示前10個非數字內容
                        if sum(1 for _ in range(10)) >= 10:
                            break
                
                if non_numeric_found and sum(1 for _ in range(10)) >= 10:
                    break
            
            if not non_numeric_found:
                print("  所有內容都是數字或允許的空值")
            
            # 檢查缺失值比例
            total_cells = df.size
            missing_cells = df.isna().sum().sum() + (df == '').sum().sum()
            missing_ratio = missing_cells / total_cells if total_cells > 0 else 0
            print(f"\n缺失值比例: {missing_ratio:.3f} ({missing_cells}/{total_cells})")
            
    except Exception as e:
        print(f"讀取檔案失敗: {e}")

def main():
    """主函數"""
    # 診斷測試檔案
    for filename in ['correct_simple.xlsx', 'defect_simple.xlsx']:
        filepath = Path(__file__).parent.parent.parent / filename
        
        if filepath.exists():
            debug_excel_file(filepath)
        else:
            print(f"檔案不存在: {filepath}")

if __name__ == '__main__':
    main()
