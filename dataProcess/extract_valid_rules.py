#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
從observed.json中提取驗證有效的數學規則
"""

import json
import os
from pathlib import Path
from typing import List, Dict, Any
import argparse
from datetime import datetime

def extract_valid_rules_from_observed(observed_file: str) -> List[str]:
    """
    從observed.json中提取所有驗證有效的規則
    
    Args:
        observed_file: observed.json檔案路徑
        
    Returns:
        List[str]: 有效規則的equation列表
    """
    valid_rules = []
    
    try:
        with open(observed_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        print(f"載入observed.json檔案: {observed_file}")
        
        # 遍歷所有執行記錄
        for script_path, execution_data in data.items():
            if 'detailed_analysis' not in execution_data:
                continue
                
            detailed_analysis = execution_data['detailed_analysis']
            print(f"\n處理執行記錄: {script_path}")
            print(f"分析時間: {execution_data.get('execution_time', 'N/A')}")
            
            # 遍歷每個窗格的分析結果
            for window_result in detailed_analysis:
                if 'validation' not in window_result:
                    continue
                
                validation = window_result['validation']
                validation_details = validation.get('validation_details', [])
                
                # 檢查每個規則的驗證結果
                for detail in validation_details:
                    if detail.get('is_valid', False):  # 程式驗證為有效
                        rule = detail.get('rule', {})
                        equation = rule.get('equation', '')
                        
                        if equation and equation not in valid_rules:
                            valid_rules.append(equation)
                            print(f"  找到有效規則: {equation}")
        
        print(f"\n總共找到 {len(valid_rules)} 個唯一的有效規則")
        return valid_rules
        
    except Exception as e:
        print(f"處理observed.json時發生錯誤: {e}")
        return []

def extract_valid_rules_by_center_row(observed_file: str) -> Dict[str, List[str]]:
    """
    按center_row分組提取有效規則
    
    Args:
        observed_file: observed.json檔案路徑
        
    Returns:
        Dict[str, List[str]]: 按center_row分組的有效規則
    """
    rules_by_center_row = {}
    
    try:
        with open(observed_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        print(f"載入observed.json檔案: {observed_file}")
        
        # 遍歷所有執行記錄
        for script_path, execution_data in data.items():
            if 'detailed_analysis' not in execution_data:
                continue
                
            detailed_analysis = execution_data['detailed_analysis']
            print(f"\n處理執行記錄: {script_path}")
            
            # 遍歷每個窗格的分析結果
            for window_result in detailed_analysis:
                if 'validation' not in window_result or 'window_info' not in window_result:
                    continue
                
                # 獲取center_row_name
                center_row_name = window_result['window_info'].get('center_row_name', 'unknown')
                
                if center_row_name not in rules_by_center_row:
                    rules_by_center_row[center_row_name] = []
                
                validation = window_result['validation']
                validation_details = validation.get('validation_details', [])
                
                # 檢查每個規則的驗證結果
                for detail in validation_details:
                    if detail.get('is_valid', False):  # 程式驗證為有效
                        rule = detail.get('rule', {})
                        equation = rule.get('equation', '')
                        
                        if equation and equation not in rules_by_center_row[center_row_name]:
                            rules_by_center_row[center_row_name].append(equation)
                            print(f"  {center_row_name}: {equation}")
        
        # 顯示統計
        print(f"\n=== 按Center Row分組統計 ===")
        for center_row, rules in rules_by_center_row.items():
            print(f"{center_row}: {len(rules)} 個有效規則")
        
        return rules_by_center_row
        
    except Exception as e:
        print(f"處理observed.json時發生錯誤: {e}")
        return {}

def save_valid_rules_json(valid_rules: List[str], output_file: str):
    """
    保存有效規則到JSON檔案
    
    Args:
        valid_rules: 有效規則列表
        output_file: 輸出檔案路徑
    """
    output_data = {
        "extraction_time": datetime.now().isoformat(),
        "total_valid_rules": len(valid_rules),
        "valid_rules": valid_rules
    }
    
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        
        print(f"有效規則已保存到: {output_file}")
        
    except Exception as e:
        print(f"保存檔案時發生錯誤: {e}")

def save_rules_by_center_row_json(rules_by_center_row: Dict[str, List[str]], output_file: str):
    """
    保存按center_row分組的有效規則到JSON檔案
    
    Args:
        rules_by_center_row: 按center_row分組的規則
        output_file: 輸出檔案路徑
    """
    output_data = {
        "extraction_time": datetime.now().isoformat(),
        "center_rows": list(rules_by_center_row.keys()),
        "total_center_rows": len(rules_by_center_row),
        "rules_by_center_row": rules_by_center_row,
        "summary": {
            center_row: len(rules) 
            for center_row, rules in rules_by_center_row.items()
        }
    }
    
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        
        print(f"按center_row分組的有效規則已保存到: {output_file}")
        
    except Exception as e:
        print(f"保存檔案時發生錯誤: {e}")

def main():
    """主函數"""
    parser = argparse.ArgumentParser(description='從observed.json中提取驗證有效的數學規則')
    parser.add_argument('--observed-file', default='../prompts/observed.json',
                       help='observed.json檔案路徑（預設: ../prompts/observed.json）')
    parser.add_argument('--output', default='valid_rules.json',
                       help='輸出檔案名稱（預設: valid_rules.json）')
    parser.add_argument('--group-by-center-row', action='store_true',
                       help='按center_row分組輸出')
    parser.add_argument('--show-details', action='store_true',
                       help='顯示詳細的規則資訊')
    
    args = parser.parse_args()
    
    # 處理檔案路徑
    observed_file = Path(__file__).parent / args.observed_file
    if not observed_file.exists():
        print(f"找不到observed.json檔案: {observed_file}")
        return
    
    output_file = Path(__file__).parent / args.output
    
    print("=== 從observed.json提取有效規則 ===")
    
    if args.group_by_center_row:
        # 按center_row分組提取
        rules_by_center_row = extract_valid_rules_by_center_row(str(observed_file))
        
        if rules_by_center_row:
            # 修改輸出檔案名稱
            output_file = output_file.with_name(f"valid_rules_by_center_row.json")
            save_rules_by_center_row_json(rules_by_center_row, str(output_file))
            
            if args.show_details:
                print(f"\n=== 詳細規則列表 ===")
                for center_row, rules in rules_by_center_row.items():
                    print(f"\n{center_row} ({len(rules)} 個規則):")
                    for i, rule in enumerate(rules, 1):
                        print(f"  {i}. {rule}")
        else:
            print("沒有找到有效規則")
    else:
        # 提取所有有效規則
        valid_rules = extract_valid_rules_from_observed(str(observed_file))
        
        if valid_rules:
            save_valid_rules_json(valid_rules, str(output_file))
            
            if args.show_details:
                print(f"\n=== 所有有效規則 ===")
                for i, rule in enumerate(valid_rules, 1):
                    print(f"{i}. {rule}")
        else:
            print("沒有找到有效規則")

if __name__ == "__main__":
    main()



