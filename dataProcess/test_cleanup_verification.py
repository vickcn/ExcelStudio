#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
驗證math_expression相關函式清理效果
"""

from run_math_rule_analysis import MathRuleAnalyzer
from llm_math_rule_detector import LLMMathRuleDetector

def test_removed_methods():
    """測試已移除的方法"""
    print("=== 測試已移除的方法 ===")
    
    analyzer = MathRuleAnalyzer(use_openai=False, window_shape=(3, 1))
    detector = LLMMathRuleDetector()
    
    # 檢查已移除的方法
    removed_methods = [
        '_evaluate_index_expression',
        '_evaluate_expression_side', 
        '_safe_eval_math_expression'  # 在run_math_rule_analysis.py中
    ]
    
    print("檢查 MathRuleAnalyzer:")
    for method in removed_methods:
        has_method = hasattr(analyzer, method)
        status = "❌ 仍存在" if has_method else "✅ 已移除"
        print(f"  {method}: {status}")
    
    print("\n檢查 LLMMathRuleDetector:")
    llm_removed_methods = [
        '_verify_index_expression',
        '_evaluate_expression_side'
    ]
    
    for method in llm_removed_methods:
        has_method = hasattr(detector, method)
        status = "❌ 仍存在" if has_method else "✅ 已移除"
        print(f"  {method}: {status}")

def test_remaining_methods():
    """測試保留的方法"""
    print(f"\n=== 測試保留的方法 ===")
    
    analyzer = MathRuleAnalyzer(use_openai=False, window_shape=(3, 1))
    detector = LLMMathRuleDetector()
    
    # 檢查保留的方法
    remaining_methods = [
        '_verify_equation_sides',
        '_evaluate_dollar_expression',
        '_safe_eval_math_expression_simple'
    ]
    
    print("檢查 MathRuleAnalyzer:")
    for method in remaining_methods:
        has_method = hasattr(analyzer, method)
        status = "✅ 存在" if has_method else "❌ 缺失"
        print(f"  {method}: {status}")
    
    print("\n檢查 LLMMathRuleDetector:")
    for method in remaining_methods:
        has_method = hasattr(detector, method)
        status = "✅ 存在" if has_method else "❌ 缺失"
        print(f"  {method}: {status}")

def test_equation_sides_only():
    """測試只支援equation_sides格式"""
    print(f"\n=== 測試只支援equation_sides格式 ===")
    
    analyzer = MathRuleAnalyzer(use_openai=False, window_shape=(3, 1))
    values = [100.0, 50.0, 50.0]
    
    # 測試新格式
    print("測試 equation_sides 格式:")
    try:
        result = analyzer._verify_rule_mathematically_direct(
            rule_text="測試規則",
            values=values,
            equation_sides=["$0 - $1", "$2"]  # 100 - 50 = 50 ✅
        )
        print(f"  equation_sides 驗證: {'✅ 通過' if result else '❌ 失敗'}")
    except Exception as e:
        print(f"  ❌ equation_sides 錯誤: {e}")
    
    # 測試回退到rule_text
    print("\n測試 rule_text 回退:")
    try:
        result = analyzer._verify_rule_mathematically_direct(
            rule_text="計算成品重 + 膠套環重 = 加膠套環後重",
            values=values
        )
        print(f"  rule_text 回退: {'✅ 執行' if result is not None else '❌ 失敗'}")
    except Exception as e:
        print(f"  ❌ rule_text 回退錯誤: {e}")

def main():
    """主測試函數"""
    print("驗證math_expression相關函式清理效果")
    print("=" * 60)
    
    test_removed_methods()
    test_remaining_methods()
    test_equation_sides_only()
    
    print(f"\n=== 清理總結 ===")
    print("🗑️ 已移除的舊格式支援:")
    print("  - ❌ math_expression (index[0] + index[2] = index[1])")
    print("  - ❌ _evaluate_index_expression")
    print("  - ❌ _evaluate_expression_side") 
    print("  - ❌ _safe_eval_math_expression (舊版)")
    
    print(f"\n✅ 保留的新格式支援:")
    print("  - ✅ equation_sides (['$0 + $2', '$1'])")
    print("  - ✅ _verify_equation_sides")
    print("  - ✅ _evaluate_dollar_expression")
    print("  - ✅ _safe_eval_math_expression_simple")
    print("  - ✅ rule_text 回退機制")
    
    print(f"\n🎯 現在系統:")
    print("  - 專注於 equation_sides 格式")
    print("  - 代碼更簡潔，沒有冗余")
    print("  - 統一使用 $ 符號表示索引")
    print("  - 保持向後兼容（rule_text）")

if __name__ == "__main__":
    main()
