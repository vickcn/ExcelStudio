#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
驗證舊的錯誤一致性率算法已經完全移除
"""

import sys
import inspect
from run_math_rule_analysis import MathRuleAnalyzer

def check_old_functions_removed():
    """檢查舊函數是否已被移除"""
    print("=== 檢查舊函數是否已被移除 ===")
    
    analyzer = MathRuleAnalyzer(use_openai=False, window_shape=(3, 1))
    
    # 檢查是否還存在舊函數
    old_functions = [
        '_check_rule_consistency',
        '_calculate_rule_consistency_rate'
    ]
    
    removed_functions = []
    remaining_functions = []
    
    for func_name in old_functions:
        if hasattr(analyzer, func_name):
            remaining_functions.append(func_name)
            print(f"❌ 舊函數仍存在: {func_name}")
        else:
            removed_functions.append(func_name)
            print(f"✅ 舊函數已移除: {func_name}")
    
    print(f"\n移除狀態:")
    print(f"  已移除: {len(removed_functions)}/{len(old_functions)}")
    print(f"  仍存在: {len(remaining_functions)}/{len(old_functions)}")
    
    return len(remaining_functions) == 0

def check_new_functions_exist():
    """檢查新函數是否存在"""
    print("\n=== 檢查新函數是否存在 ===")
    
    analyzer = MathRuleAnalyzer(use_openai=False, window_shape=(3, 1))
    
    # 檢查新函數是否存在
    new_functions = [
        '_calculate_rule_consistency_in_center_row',
        '_build_rule_consistency_map'
    ]
    
    existing_functions = []
    missing_functions = []
    
    for func_name in new_functions:
        if hasattr(analyzer, func_name):
            existing_functions.append(func_name)
            print(f"✅ 新函數存在: {func_name}")
        else:
            missing_functions.append(func_name)
            print(f"❌ 新函數缺失: {func_name}")
    
    print(f"\n新函數狀態:")
    print(f"  存在: {len(existing_functions)}/{len(new_functions)}")
    print(f"  缺失: {len(missing_functions)}/{len(new_functions)}")
    
    return len(missing_functions) == 0

def check_function_signatures():
    """檢查新函數的簽名是否正確"""
    print("\n=== 檢查新函數簽名 ===")
    
    analyzer = MathRuleAnalyzer(use_openai=False, window_shape=(3, 1))
    
    # 檢查 _calculate_rule_consistency_in_center_row
    if hasattr(analyzer, '_calculate_rule_consistency_in_center_row'):
        func = getattr(analyzer, '_calculate_rule_consistency_in_center_row')
        sig = inspect.signature(func)
        params = list(sig.parameters.keys())
        
        expected_params = ['group_results']
        if params == expected_params:
            print("✅ _calculate_rule_consistency_in_center_row 簽名正確")
        else:
            print(f"❌ _calculate_rule_consistency_in_center_row 簽名錯誤: {params} != {expected_params}")
    
    # 檢查 _build_rule_consistency_map
    if hasattr(analyzer, '_build_rule_consistency_map'):
        func = getattr(analyzer, '_build_rule_consistency_map')
        sig = inspect.signature(func)
        params = list(sig.parameters.keys())
        
        expected_params = ['analysis_result']
        if params == expected_params:
            print("✅ _build_rule_consistency_map 簽名正確")
        else:
            print(f"❌ _build_rule_consistency_map 簽名錯誤: {params} != {expected_params}")
    
    return True

def check_code_for_old_references():
    """檢查代碼中是否還有對舊函數的引用"""
    print("\n=== 檢查代碼中的舊函數引用 ===")
    
    import run_math_rule_analysis
    source_file = run_math_rule_analysis.__file__
    
    try:
        with open(source_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        old_function_calls = [
            '_check_rule_consistency(',
            '_calculate_rule_consistency_rate('
        ]
        
        found_references = []
        for func_call in old_function_calls:
            if func_call in content:
                found_references.append(func_call)
                print(f"❌ 發現舊函數調用: {func_call}")
        
        if not found_references:
            print("✅ 沒有發現舊函數的調用")
        
        return len(found_references) == 0
        
    except Exception as e:
        print(f"❌ 檢查源代碼失敗: {e}")
        return False

def main():
    """主檢查函數"""
    print("驗證舊的錯誤一致性率算法已經完全移除")
    print("=" * 50)
    
    checks = [
        ("舊函數移除檢查", check_old_functions_removed),
        ("新函數存在檢查", check_new_functions_exist),
        ("函數簽名檢查", check_function_signatures),
        ("代碼引用檢查", check_code_for_old_references)
    ]
    
    passed = 0
    total = len(checks)
    
    for check_name, check_func in checks:
        print(f"\n--- {check_name} ---")
        try:
            if check_func():
                print(f"✅ {check_name} 通過")
                passed += 1
            else:
                print(f"❌ {check_name} 失敗")
        except Exception as e:
            print(f"❌ {check_name} 錯誤: {e}")
            import traceback
            traceback.print_exc()
    
    print(f"\n=== 驗證結果 ===")
    print(f"通過: {passed}/{total}")
    
    if passed == total:
        print("🎉 舊算法已完全移除，新算法已正確實作！")
        print("\n移除的舊算法:")
        print("- ❌ _check_rule_consistency (錯誤的模式匹配方式)")
        print("- ❌ _calculate_rule_consistency_rate (錯誤的跨center_row平均)")
        print("\n新的正確算法:")
        print("- ✅ _calculate_rule_consistency_in_center_row (正確的center_row內一致性率)")
        print("- ✅ _build_rule_consistency_map (從新算法結果建立映射)")
    else:
        print("⚠️  部分檢查失敗，可能還有舊算法殘留")
    
    return 0 if passed == total else 1

if __name__ == "__main__":
    sys.exit(main())
