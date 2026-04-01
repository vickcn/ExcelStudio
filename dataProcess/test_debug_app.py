#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試整合偵錯應用程式
"""

import os
import sys
import json
from pathlib import Path

# 添加父目錄到路徑
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from integrated_debug_app import IntegratedDebugApp

def test_debug_app():
    """測試偵錯應用程式"""
    print("=== 測試整合偵錯應用程式 ===")
    
    # 測試文件路徑
    test_files = [
        "D:/JobProject/TWSG/CE/重量規格/correct_simple.xlsx",
        "D:/JobProject/TWSG/CE/重量規格/defect_simple.xlsx"
    ]
    
    # 檢查測試文件是否存在
    available_files = []
    for file_path in test_files:
        if os.path.exists(file_path):
            available_files.append(file_path)
            print(f"✓ 找到測試文件: {file_path}")
        else:
            print(f"✗ 測試文件不存在: {file_path}")
    
    if not available_files:
        print("錯誤: 沒有可用的測試文件")
        return
    
    # 初始化偵錯應用程式
    try:
        debug_app = IntegratedDebugApp()
        print("✓ 偵錯應用程式初始化成功")
    except Exception as e:
        print(f"✗ 偵錯應用程式初始化失敗: {e}")
        return
    
    # 測試每個可用文件
    for test_file in available_files:
        print(f"\n--- 測試文件: {os.path.basename(test_file)} ---")
        
        try:
            # 執行偵錯分析
            results = debug_app.run_debug_analysis(
                test_file, 
                use_openai=False,  # 使用本地LLM
                output_dir=f"./test_results_{os.path.basename(test_file).replace('.xlsx', '')}"
            )
            
            if 'error' in results:
                print(f"✗ 分析失敗: {results['error']}")
            else:
                summary = results['summary']
                print(f"✓ 分析完成")
                print(f"  - 偵測到數據範圍: {results['data_ranges_found']}")
                print(f"  - 總驗證次數: {summary['total_validations']}")
                print(f"  - 通過驗證: {summary['passed_validations']}")
                print(f"  - 失敗驗證: {summary['failed_validations']}")
                print(f"  - 發現異常: {summary['anomalies_count']}")
                print(f"  - 成功率: {summary['success_rate']:.2%}")
                
                # 顯示異常詳情（如果有）
                if summary['anomalies_count'] > 0:
                    print("  異常位置:")
                    for i, anomaly in enumerate(results['validation_results']['anomalies'][:3], 1):  # 只顯示前3個
                        location = anomaly['location']
                        rule_desc = anomaly.get('rule_description', 'Unknown rule')
                        print(f"    {i}. 位置 {location}: {rule_desc}")
                    
                    if summary['anomalies_count'] > 3:
                        print(f"    ... 還有 {summary['anomalies_count'] - 3} 個異常")
        
        except Exception as e:
            print(f"✗ 測試失敗: {e}")
            import traceback
            traceback.print_exc()
    
    print("\n=== 測試完成 ===")

def test_individual_components():
    """測試個別組件"""
    print("\n=== 測試個別組件 ===")
    
    try:
        debug_app = IntegratedDebugApp()
        
        # 測試規則載入
        print(f"載入的規則數量: {len(debug_app.discovered_rules)}")
        for rule_key, rule_info in list(debug_app.discovered_rules.items())[:2]:  # 只顯示前2個
            print(f"  規則 {rule_key}: {rule_info.get('start_loc_row_name', 'Unknown')}")
        
        # 測試文件存在性檢查
        test_file = "D:/JobProject/TWSG/CE/重量規格/correct_simple.xlsx"
        if os.path.exists(test_file):
            import pandas as pd
            df = pd.read_excel(test_file)
            print(f"✓ 成功載入測試文件，形狀: {df.shape}")
            
            # 測試數據範圍偵測
            ranges = debug_app.detect_data_ranges(df)
            print(f"✓ 偵測到 {len(ranges)} 個數據範圍")
            
            for i, range_info in enumerate(ranges[:2], 1):  # 只顯示前2個
                print(f"  範圍 {i}: 行 {range_info['start_row']}-{range_info['end_row']}, 列 {range_info['start_col']}-{range_info['end_col']}")
                print(f"    行名稱: {range_info['row_names'][:3]}...")  # 只顯示前3個行名稱
        
    except Exception as e:
        print(f"✗ 組件測試失敗: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_individual_components()
    test_debug_app()


