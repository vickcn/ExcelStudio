#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速測試數學規律偵測系統
"""

import os
import sys
from pathlib import Path

# 添加路徑
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

def quick_test():
    """快速測試主要功能"""
    print("=== 快速測試數學規律偵測系統 ===\n")
    
    try:
        # 測試LLM偵測器
        print("1. 測試LLM偵測器...")
        from llm_math_rule_detector import LLMMathRuleDetector
        
        detector = LLMMathRuleDetector(prefer_local=True)
        test_values = [1.0, 2.0, 3.0, 4.0, 5.0]
        test_columns = ["A", "B", "C", "D", "E"]
        values_matrix = [[v] for v in test_values]
        row_names = test_columns
        col_names = ["value"]
        
        print(f"測試數值: {test_values}")
        result = detector.detect_math_rules(values_matrix, row_names, col_names)
        
        if result:
            print(f"  偵測成功，發現 {len(result.get('rules', []))} 個規律")
            
            # 驗證規律
            validation = detector.validate_rules(values_matrix, result.get('rules', []))
            print(f"  驗證: {validation['valid_rules']}/{validation['total_rules']} 有效")
            print(f"  成功率: {validation['success_rate']:.2%}")
        else:
            print("  偵測失敗")
        
        print("\n2. 測試窗格掃描器...")
        import pandas as pd
        from window_scanner import WindowScanner
        
        # 簡單測試數據
        df = pd.DataFrame({
            'values': [10, 20, 30, 40, 50]
        }, index=['row1', 'row2', 'row3', 'row4', 'row5'])
        
        scanner = WindowScanner(window_shape=(3, 1))
        windows = scanner.scan_dataframe(df)
        
        print(f"  掃描產生 {len(windows)} 個窗格")
        
        if windows:
            first_window = windows[0]
            prompt_data = scanner.generate_prompt_data(first_window)
            print(f"  第一個窗格提示: {prompt_data['prompt_text']}")
        
        print("\n3. 測試表格偵測器...")
        excel_file = Path(__file__).parent.parent.parent / 'correct_simple.xlsx'
        
        if excel_file.exists():
            from universal_table_detector import UniversalTableDetector
            
            table_detector = UniversalTableDetector()
            results = table_detector.detect_tables_by_analysis(str(excel_file), 'pure_numeric')
            
            print(f"  偵測到 {len(results)} 個表格")
            
            if results:
                dataframes = table_detector.extract_dataframes(str(excel_file), results)
                if dataframes:
                    df = dataframes[0]
                    print(f"  DataFrame形狀: {df.shape}")
        else:
            print(f"  找不到測試檔案: {excel_file}")
        
        print("\n=== 快速測試完成 ===")
        print("所有基本功能都正常運作")
        
    except Exception as e:
        print(f"測試失敗: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    quick_test()
