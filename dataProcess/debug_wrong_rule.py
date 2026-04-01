#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
調試錯誤規則的驗證問題
"""

from run_math_rule_analysis import MathRuleAnalyzer

def test_wrong_rule_validation():
    """測試錯誤規則的驗證"""
    print("=== 測試錯誤規則驗證 ===")
    
    analyzer = MathRuleAnalyzer(use_openai=False, window_shape=(3, 1))
    
    # 模擬典型的重量數據
    test_cases = [
        {
            "name": "典型數據1",
            "values": [237.6, 238.8, 1.2],  # 計算成品重, 加膠套環後重, 膠套環重
            "description": "計算成品重: 237.6, 加膠套環後重: 238.8, 膠套環重: 1.2"
        },
        {
            "name": "典型數據2", 
            "values": [100.0, 110.0, 10.0],
            "description": "計算成品重: 100.0, 加膠套環後重: 110.0, 膠套環重: 10.0"
        }
    ]
    
    # 測試的規則
    rules_to_test = [
        {
            "rule": "計算成品重（不含膠套環）減去膠套環重等於加膠套環後的重量",
            "expected": False,
            "math": "計算成品重 - 膠套環重 = 加膠套環後重"
        },
        {
            "rule": "計算成品重+膠套環重=加膠套環後重",
            "expected": True,
            "math": "計算成品重 + 膠套環重 = 加膠套環後重"
        },
        {
            "rule": "加膠套環後重=計算成品重+膠套環重",
            "expected": True,
            "math": "加膠套環後重 = 計算成品重 + 膠套環重"
        }
    ]
    
    for test_case in test_cases:
        print(f"\n--- {test_case['name']} ---")
        print(f"數據: {test_case['description']}")
        values = test_case['values']
        
        計算成品重, 加膠套環後重, 膠套環重 = values
        
        print(f"實際關係驗證:")
        print(f"  計算成品重 + 膠套環重 = {計算成品重} + {膠套環重} = {計算成品重 + 膠套環重}")
        print(f"  加膠套環後重 = {加膠套環後重}")
        print(f"  正確關係: {abs((計算成品重 + 膠套環重) - 加膠套環後重) < 0.01}")
        
        print(f"  計算成品重 - 膠套環重 = {計算成品重} - {膠套環重} = {計算成品重 - 膠套環重}")
        print(f"  錯誤關係: {abs((計算成品重 - 膠套環重) - 加膠套環後重) < 0.01}")
        
        for rule_test in rules_to_test:
            print(f"\n  測試規則: {rule_test['rule']}")
            print(f"  數學表達: {rule_test['math']}")
            
            try:
                result = analyzer._verify_rule_mathematically_direct(
                    rule_text=rule_test['rule'],
                    values=values
                )
                
                expected = rule_test['expected']
                status = "✅" if result == expected else "❌"
                
                print(f"  驗證結果: {result} (期望: {expected}) {status}")
                
                if result != expected:
                    print(f"  ⚠️ 驗證結果不符合期望！")
                    
            except Exception as e:
                print(f"  ❌ 驗證異常: {e}")

def test_rule_text_matching():
    """測試規則文字匹配"""
    print(f"\n=== 測試規則文字匹配 ===")
    
    analyzer = MathRuleAnalyzer(use_openai=False, window_shape=(3, 1))
    
    # 測試規則文字正規化
    test_rules = [
        "計算成品重（不含膠套環）減去膠套環重等於加膠套環後的重量",
        "計算成品重-膠套環重=加膠套環後重",
        "計算成品重 - 膠套環重 = 加膠套環後重"
    ]
    
    for rule in test_rules:
        normalized = rule.replace(' ', '').replace('\n', '')
        print(f"原始: {rule}")
        print(f"正規化: {normalized}")
        
        # 檢查匹配條件
        matches = []
        if '計算成品重+膠套環重=加膠套環後重' in normalized:
            matches.append("正確加法關係")
        if '計算成品重-膠套環重=加膠套環後重' in normalized:
            matches.append("錯誤減法關係")
        if '加膠套環後重=計算成品重+膠套環重' in normalized:
            matches.append("正確等式關係")
            
        print(f"匹配到: {matches if matches else '無匹配'}")
        print()

def main():
    """主測試函數"""
    print("調試錯誤規則的驗證問題")
    print("=" * 60)
    
    test_wrong_rule_validation()
    test_rule_text_matching()
    
    print(f"\n=== 問題分析 ===")
    print("可能的問題原因:")
    print("1. 🔍 規則文字匹配邏輯問題")
    print("2. 📊 實際數據可能真的滿足錯誤關係") 
    print("3. 🤖 LLM生成了錯誤的規則描述")
    print("4. 🔧 驗證邏輯中的數學計算錯誤")
    
    print(f"\n💡 建議檢查:")
    print("- 檢查實際的數據值")
    print("- 驗證規則文字正規化邏輯")
    print("- 確認數學計算的正確性")
    print("- 檢查LLM生成的原始規則")

if __name__ == "__main__":
    main()
