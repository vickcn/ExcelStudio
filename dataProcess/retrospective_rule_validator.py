#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
回溯式規則驗證器
當發現新規則時，回溯驗證同一center_row的所有窗格
"""

from collections import defaultdict
from typing import Dict, List, Any
from llm_math_rule_detector import LLMMathRuleDetector

class RetrospectiveRuleValidator:
    """回溯式規則驗證器"""
    
    def __init__(self):
        self.llm_detector = LLMMathRuleDetector()
    
    def validate_rules_retrospectively(self, center_row_groups: Dict[str, List[Dict]]) -> Dict[str, Any]:
        """
        對每個center_row進行回溯式規則驗證
        
        Args:
            center_row_groups: 按center_row分組的窗格結果
            
        Returns:
            Dict: 回溯驗證後的結果
        """
        retrospective_results = {}
        
        for center_row, group_results in center_row_groups.items():
            print(f"\n=== 回溯驗證 Center Row: {center_row} ===")
            
            # 第一步：收集所有規則
            all_rules = self._collect_all_rules_in_center_row(group_results)
            print(f"發現 {len(all_rules)} 個不同的規則")
            
            # 第二步：對每個窗格驗證所有規則
            retrospective_validation = self._validate_all_rules_in_all_windows(
                group_results, all_rules
            )
            
            # 第三步：計算真實的一致性率
            rule_consistency = self._calculate_true_consistency_rates(
                retrospective_validation, len(group_results)
            )
            
            retrospective_results[center_row] = {
                'window_count': len(group_results),
                'total_rules': len(all_rules),
                'rule_consistency': rule_consistency,
                'retrospective_validation': retrospective_validation
            }
            
            # 顯示結果
            for rule_text, stats in rule_consistency.items():
                print(f"  規則: {rule_text[:60]}{'...' if len(rule_text) > 60 else ''}")
                print(f"    一致性率: {stats['consistency_rate']:.1%} ({stats['valid_windows']}/{stats['total_windows']})")
        
        return retrospective_results
    
    def _collect_all_rules_in_center_row(self, group_results: List[Dict]) -> List[str]:
        """收集center_row中所有出現過的規則"""
        all_rules = set()
        
        for window_result in group_results:
            if 'llm_result' in window_result and 'rules' in window_result['llm_result']:
                for rule in window_result['llm_result']['rules']:
                    rule_text = rule.get('equation', rule.get('rule', rule.get('description', '')))
                    if rule_text:
                        normalized_rule = ' '.join(rule_text.split()).strip()
                        all_rules.add(normalized_rule)
        
        return list(all_rules)
    
    def _validate_all_rules_in_all_windows(self, group_results: List[Dict], all_rules: List[str]) -> Dict[str, Dict]:
        """在所有窗格中驗證所有規則"""
        validation_matrix = {}
        
        for window_result in group_results:
            window_id = window_result.get('window_info', {}).get('window_id', 'unknown')
            window_values = self._extract_window_values(window_result)
            
            validation_matrix[window_id] = {}
            
            for rule_text in all_rules:
                # 為每個規則創建驗證用的規則對象
                rule_obj = {
                    'equation': rule_text,
                    'equationWithValues': self._generate_equation_with_values(rule_text, window_values),
                    'confidence': 1.0
                }
                
                # 驗證規則
                validation_result = self.llm_detector.validate_rules(window_values, [rule_obj])
                
                is_valid = False
                if validation_result['validation_details']:
                    is_valid = validation_result['validation_details'][0].get('is_valid', False)
                
                validation_matrix[window_id][rule_text] = {
                    'is_valid': is_valid,
                    'validation_details': validation_result['validation_details'][0] if validation_result['validation_details'] else {}
                }
        
        return validation_matrix
    
    def _extract_window_values(self, window_result: Dict) -> List[List[Any]]:
        """從窗格結果中提取數值矩陣"""
        if 'prompt_data' in window_result and 'values' in window_result['prompt_data']:
            values = window_result['prompt_data']['values']
            if isinstance(values, list) and values and isinstance(values[0], list):
                return values
            if isinstance(values, list):
                return [[v] for v in values]
        elif 'window_info' in window_result and 'values' in window_result['window_info']:
            values = window_result['window_info']['values']
            if isinstance(values, list) and values:
                if isinstance(values[0], list):
                    return values
                return [[v] for v in values]
        return []
    
    def _generate_equation_with_values(self, rule_text: str, values: List[List[Any]]) -> str:
        """為規則生成帶數值的等式（簡化版）"""
        if values and len(values) >= 3 and values[0] and len(values[0]) >= 1:
            return f"{values[0][0]} + {values[1][0]} = {values[2][0]}"
        return rule_text
    
    def _calculate_true_consistency_rates(self, validation_matrix: Dict, total_windows: int) -> Dict[str, Dict]:
        """計算真實的一致性率"""
        rule_stats = defaultdict(lambda: {
            'total_windows': total_windows,
            'valid_windows': 0,
            'invalid_windows': 0,
            'consistency_rate': 0.0,
            'validation_details': []
        })
        
        # 統計每個規則在所有窗格中的驗證結果
        for window_id, window_validations in validation_matrix.items():
            for rule_text, validation_info in window_validations.items():
                is_valid = validation_info['is_valid']
                
                if is_valid:
                    rule_stats[rule_text]['valid_windows'] += 1
                else:
                    rule_stats[rule_text]['invalid_windows'] += 1
                
                rule_stats[rule_text]['validation_details'].append({
                    'window_id': window_id,
                    'is_valid': is_valid,
                    'details': validation_info['validation_details']
                })
        
        # 計算一致性率
        for rule_text in rule_stats:
            valid_count = rule_stats[rule_text]['valid_windows']
            rule_stats[rule_text]['consistency_rate'] = valid_count / total_windows
        
        return dict(rule_stats)

def main():
    """測試回溯式驗證"""
    print("回溯式規則驗證器測試")
    print("=" * 50)
    
    # 這裡可以添加測試代碼
    validator = RetrospectiveRuleValidator()
    print("✅ 回溯式驗證器初始化成功")
    
    return 0

if __name__ == "__main__":
    import sys
    sys.exit(main())
