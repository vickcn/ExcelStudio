#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
調試$符號表達式評估
"""

from llm_math_rule_detector import LLMMathRuleDetector

def test_dollar_expression():
    """測試$符號表達式評估"""
    print("=== 測試$符號表達式評估 ===")
    
    detector = LLMMathRuleDetector()
    
    # 使用測試數據
    values = [238.8, 237.6, 238.8, 1.2, 237.6]
    print(f"測試數據: {values}")
    
    # 測試各種表達式
    test_cases = [
        "$0",           # 238.8
        "$1",           # 237.6
        "$0 - $1",      # 238.8 - 237.6 = 1.2
        "$0 + $2",      # 238.8 + 238.8 = 477.6
        "$0 - $3 + $2", # 238.8 - 1.2 + 238.8 = 476.4
    ]
    
    for expr in test_cases:
        try:
            result = detector._evaluate_dollar_expression(expr, values)
            print(f"表達式: {expr} = {result}")
        except Exception as e:
            print(f"表達式: {expr} 錯誤: {e}")

def test_equation_sides_validation():
    """測試equation_sides驗證"""
    print("\n=== 測試equation_sides驗證 ===")
    
    detector = LLMMathRuleDetector()
    
    # 使用測試數據
    values = [238.8, 237.6, 238.8, 1.2, 237.6]
    print(f"測試數據: {values}")
    
    # 測試OpenAI回傳的規則
    test_rules = [
        ["$0 - $1", "$1"],        # 238.8 - 237.6 = 1.2, 但右邊是 237.6
        ["$0 + $2", "$1"],        # 238.8 + 238.8 = 477.6, 但右邊是 237.6  
        ["$0 - $3 + $2", "$4"],   # 238.8 - 1.2 + 238.8 = 476.4, 但右邊是 237.6
    ]
    
    for i, rule in enumerate(test_rules):
        print(f"\n規則 {i+1}: {rule}")
        
        try:
            # 計算左邊
            left_value = detector._evaluate_dollar_expression(rule[0], values)
            print(f"  左邊 {rule[0]} = {left_value}")
            
            # 計算右邊
            right_value = detector._evaluate_dollar_expression(rule[1], values)
            print(f"  右邊 {rule[1]} = {right_value}")
            
            # 檢查是否相等
            is_equal = abs(left_value - right_value) < 0.01
            print(f"  相等: {is_equal} (差異: {abs(left_value - right_value):.6f})")
            
            # 使用系統驗證
            system_result = detector._verify_equation_sides(rule, values)
            print(f"  系統驗證: {system_result}")
            
        except Exception as e:
            print(f"  錯誤: {e}")

def test_correct_rules():
    """測試正確的規則"""
    print("\n=== 測試正確的規則 ===")
    
    detector = LLMMathRuleDetector()
    
    # 使用測試數據
    values = [238.8, 237.6, 238.8, 1.2, 237.6]
    print(f"測試數據: {values}")
    print("觀察: values[0] = values[2] = 238.8")
    print("觀察: values[1] = values[4] = 237.6") 
    print("觀察: values[0] - values[1] = 238.8 - 237.6 = 1.2 = values[3]")
    
    # 正確的規則
    correct_rules = [
        ["$0", "$2"],           # 238.8 = 238.8 ✅
        ["$1", "$4"],           # 237.6 = 237.6 ✅
        ["$0 - $1", "$3"],      # 238.8 - 237.6 = 1.2 ✅
    ]
    
    for i, rule in enumerate(correct_rules):
        print(f"\n正確規則 {i+1}: {rule}")
        
        try:
            # 計算左邊
            left_value = detector._evaluate_dollar_expression(rule[0], values)
            print(f"  左邊 {rule[0]} = {left_value}")
            
            # 計算右邊
            right_value = detector._evaluate_dollar_expression(rule[1], values)
            print(f"  右邊 {rule[1]} = {right_value}")
            
            # 檢查是否相等
            is_equal = abs(left_value - right_value) < 0.01
            print(f"  相等: {is_equal} ✅" if is_equal else f"  相等: {is_equal} ❌")
            
            # 使用系統驗證
            system_result = detector._verify_equation_sides(rule, values)
            print(f"  系統驗證: {system_result} ✅" if system_result else f"  系統驗證: {system_result} ❌")
            
        except Exception as e:
            print(f"  錯誤: {e}")

if __name__ == "__main__":
    test_dollar_expression()
    test_equation_sides_validation()
    test_correct_rules()
