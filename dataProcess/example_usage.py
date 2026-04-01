#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
數學規則偵測器使用範例
展示如何使用地端LLM和OpenAI API
"""

import os
import sys
import subprocess

def run_command(cmd, description):
    """執行命令並顯示結果"""
    print(f"\n{'='*50}")
    print(f"執行: {description}")
    print(f"命令: {cmd}")
    print(f"{'='*50}")
    
    try:
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True, encoding='utf-8')
        
        if result.stdout:
            print("輸出:")
            print(result.stdout)
        
        if result.stderr:
            print("錯誤:")
            print(result.stderr)
        
        print(f"返回碼: {result.returncode}")
        
    except Exception as e:
        print(f"執行失敗: {e}")

def main():
    """主程式"""
    print("數學規則偵測器使用範例")
    print("這個腳本展示如何使用不同的參數來運行分析")
    
    # 檢查數據文件是否存在
    data_file = "processed_window_data_20251003_000747_5_1.xlsx"
    if not os.path.exists(data_file):
        print(f"\n警告: 數據文件不存在 - {data_file}")
        print("以下命令將會失敗，但可以看到正確的使用方式")
    
    # 範例1: 使用地端LLM（預設）
    cmd1 = f"python math_rule_detector.py --file {data_file}"
    run_command(cmd1, "使用地端LLM分析數學規則")
    
    # 範例2: 使用OpenAI API
    cmd2 = f"python math_rule_detector.py --file {data_file} --use-openai"
    run_command(cmd2, "使用OpenAI API分析數學規則")
    
    # 範例3: 指定輸出文件
    cmd3 = f"python math_rule_detector.py --file {data_file} --output custom_output.json"
    run_command(cmd3, "指定自定義輸出文件")
    
    # 範例4: 指定工作表
    cmd4 = f"python math_rule_detector.py --file {data_file} --sheet Sheet1"
    run_command(cmd4, "指定特定工作表")
    
    # 範例5: 使用簡化腳本（地端LLM）
    cmd5 = "python run_math_rule_analysis.py"
    run_command(cmd5, "使用簡化腳本（預設地端LLM）")
    
    # 範例6: 使用簡化腳本（OpenAI）
    cmd6 = "python run_math_rule_analysis.py --use-openai"
    run_command(cmd6, "使用簡化腳本（OpenAI API）")
    
    print(f"\n{'='*50}")
    print("所有範例執行完成")
    print("注意：實際執行需要確保:")
    print("1. 數據文件存在")
    print("2. 地端LLM服務運行中（如果使用地端LLM）")
    print("3. OpenAI API金鑰已設定（如果使用OpenAI）")
    print(f"{'='*50}")

if __name__ == "__main__":
    main()




