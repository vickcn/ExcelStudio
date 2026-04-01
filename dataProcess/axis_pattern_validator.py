#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Axis 方向規律驗證器 - 在檢測到有規律的方向進行窗格掃描和驗證
"""

import os
import json
import pandas as pd
import numpy as np
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
import argparse
from datetime import datetime as dt
import requests

from utils import LOGger

m_logfile = Path(__file__).resolve().parent.parent / "outputs" / "axis_pattern_validator.log"
m_print = LOGger.addloger(logfile=str(m_logfile))

# 導入自定義模組
from universal_table_detector import UniversalTableDetector
from llm_math_rule_detector import LLMMathRuleDetector


class AxisPatternValidator:
    """Axis 方向規律驗證器"""
    
    def __init__(self, window_size: int = 3, use_remote: bool = True):
        """
        初始化驗證器
        
        Args:
            window_size: 窗格大小（建議3-5）
            use_remote: 是否使用 remote 8b（預設 True）
        """
        self.window_size = window_size
        
        # 從 config.json 載入設定並設置環境變數（必須在初始化 LLM 偵測器之前）
        config_path = Path(__file__).parent.parent.parent / 'config.json'
        if config_path.exists():
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                textprocessor_config = config.get('textprocessor', {})
                # 設置環境變數，讓 LLMMathRuleDetector 使用這些設置
                os.environ['TEXTPROCESSOR_URL'] = textprocessor_config.get('url', 'http://10.1.3.127:6017/chat')
                os.environ['TEXTPROCESSOR_PROVIDER'] = textprocessor_config.get('provider', 'remote')
                os.environ['TEXTPROCESSOR_MODEL'] = textprocessor_config.get('model', 'remote8b')
                m_print(f"已從 config.json 載入 TextProcessor 設定")
                m_print(f"Provider: {textprocessor_config.get('provider', 'remote')}, Model: {textprocessor_config.get('model', 'remote8b')}")
        
        # 初始化組件（順序很重要）
        self.table_detector = UniversalTableDetector()
        self.llm_detector = LLMMathRuleDetector(prefer_local=True)  # 使用 TextProcessor
        
        m_print(f"Axis 方向規律驗證器已初始化（窗格大小: {window_size}）")
    
    def validate_pattern(self, pattern_result_file: str, excel_file: str) -> Dict[str, Any]:
        """
        驗證有規律的方向
        
        Args:
            pattern_result_file: height_pattern_result.json 路徑
            excel_file: Excel檔案路徑
            
        Returns:
            Dict: 驗證結果
        """
        # 讀取 pattern 結果
        pattern_result = self._load_pattern_result(pattern_result_file)
        if not pattern_result:
            return {'success': False, 'error': '無法載入 pattern 結果'}
        
        llm_analysis = pattern_result.get('llm_analysis', {})
        
        # 確定要驗證的方向
        target_axis = None
        if llm_analysis.get('axis0_confidence', 0) > 0.5:
            target_axis = 0
            m_print(f"檢測到 axis=0 有規律，信心度: {llm_analysis.get('axis0_confidence', 0):.2%}")
        elif llm_analysis.get('axis1_confidence', 0) > 0.5:
            target_axis = 1
            m_print(f"檢測到 axis=1 有規律，信心度: {llm_analysis.get('axis1_confidence', 0):.2%}")
        else:
            return {'success': False, 'error': '沒有達到信度門檻的方向'}
        
        # 讀取 Excel 檔案
        excel_path = Path(excel_file)
        if not excel_path.is_absolute():
            base_path = Path(__file__).parent.parent.parent
            excel_path = base_path / excel_file
        
        if not excel_path.exists():
            return {'success': False, 'error': f'找不到檔案: {excel_path}'}
        
        # 偵測表格
        m_print(f"偵測表格: {excel_path}")
        table_results = self.table_detector.detect_tables_by_analysis(
            str(excel_path), 
            table_mode='pure_numeric',
            use_llm=False
        )
        
        if not table_results:
            return {'success': False, 'error': '未偵測到有效表格'}
        
        dataframes = self.table_detector.extract_dataframes(str(excel_path), table_results)
        main_df = dataframes[0]
        
        m_print(f"表格形狀: {main_df.shape}")
        
        # 根據 axis 進行窗格掃描和驗證
        if target_axis == 0:
            result = self._validate_axis0(main_df, llm_analysis)
        else:
            result = self._validate_axis1(main_df, llm_analysis)
        
        result['target_axis'] = target_axis
        result['excel_file'] = str(excel_path)
        result['analysis_time'] = dt.now().isoformat()
        
        return result
    
    def _validate_axis0(self, df: pd.DataFrame, llm_analysis: Dict) -> Dict[str, Any]:
        """
        驗證 axis=0（行方向）的規律
        
        Args:
            df: DataFrame
            llm_analysis: LLM 分析結果
            
        Returns:
            Dict: 驗證結果
        """
        m_print("=== 掃描 axis=0（行方向）===")
        
        validation_results = []
        total_windows = 0
        valid_windows = 0
        failed_cells = []
        
        # 沿行掃描
        for row_idx in range(len(df)):
            # 只在有足夠列數時才能形成窗格
            if df.shape[1] < self.window_size:
                continue
            
            for start_col in range(df.shape[1] - self.window_size + 1):
                end_col = start_col + self.window_size
                
                # 提取窗格
                window_values = []
                window_cols = []
                for col_idx in range(start_col, end_col):
                    value = df.iloc[row_idx, col_idx]
                    window_values.append(value)
                    window_cols.append(df.columns[col_idx])
                
                total_windows += 1
                
                # 只處理有效數值
                numeric_values = []
                for val in window_values:
                    try:
                        numeric_values.append(float(val))
                    except (ValueError, TypeError):
                        pass
                
                if len(numeric_values) < self.window_size:
                    failed_cells.append({
                        'axis': 0,
                        'row': row_idx,
                        'row_name': df.index[row_idx],
                        'start_col': start_col,
                        'end_col': end_col - 1,
                        'columns': window_cols,
                        'reason': '非數字值過多',
                        'values': window_values
                    })
                    continue
                
                # 使用 LLM 偵測規律
                values_matrix = [numeric_values]
                row_names = [str(df.index[row_idx])]
                col_names = [str(c) for c in window_cols]
                llm_result = self.llm_detector.detect_math_rules(
                    values_matrix,
                    row_names,
                    col_names,
                    use_openai=False
                )
                
                if llm_result and llm_result.get('rules'):
                    # 驗證規律
                    validation = self.llm_detector.validate_rules(
                        values_matrix,
                        llm_result.get('rules')
                    )
                    
                    if validation.get('valid_rules', 0) > 0:
                        valid_windows += 1
                    else:
                        failed_cells.append({
                            'axis': 0,
                            'row': row_idx,
                            'row_name': str(df.index[row_idx]),
                            'start_col': start_col,
                            'end_col': end_col - 1,
                            'columns': window_cols,
                            'reason': '規則無法驗證',
                            'values': numeric_values,
                            'rules': [r.get('equation') for r in llm_result.get('rules', [])]
                        })
                else:
                    failed_cells.append({
                        'axis': 0,
                        'row': row_idx,
                        'row_name': str(df.index[row_idx]),
                        'start_col': start_col,
                        'end_col': end_col - 1,
                        'columns': window_cols,
                        'reason': '未偵測到規律',
                        'values': numeric_values
                    })
        
        success_rate = valid_windows / total_windows if total_windows > 0 else 0
        
        m_print(f"總窗格數: {total_windows}")
        m_print(f"有效窗格數: {valid_windows}")
        m_print(f"成功率: {success_rate:.2%}")
        m_print(f"失敗位置數: {len(failed_cells)}")
        
        return {
            'success': True,
            'axis': 0,
            'description': llm_analysis.get('axis0_description', ''),
            'confidence': llm_analysis.get('axis0_confidence', 0),
            'total_windows': total_windows,
            'valid_windows': valid_windows,
            'success_rate': success_rate,
            'failed_cells': failed_cells,
            'failed_count': len(failed_cells)
        }
    
    def _validate_axis1(self, df: pd.DataFrame, llm_analysis: Dict) -> Dict[str, Any]:
        """
        驗證 axis=1（列方向）的規律
        
        Args:
            df: DataFrame
            llm_analysis: LLM 分析結果
            
        Returns:
            Dict: 驗證結果
        """
        m_print("=== 掃描 axis=1（列方向）===")
        
        total_windows = 0
        valid_windows = 0
        failed_cells = []
        
        # 沿列掃描
        for col_idx in range(df.shape[1]):
            # 只在有足夠行數時才能形成窗格
            if df.shape[0] < self.window_size:
                continue
            
            for start_row in range(df.shape[0] - self.window_size + 1):
                end_row = start_row + self.window_size
                
                # 提取窗格
                window_values = []
                window_rows = []
                for row_idx in range(start_row, end_row):
                    value = df.iloc[row_idx, col_idx]
                    window_values.append(value)
                    window_rows.append(df.index[row_idx])
                
                total_windows += 1
                
                # 只處理有效數值
                numeric_values = []
                for val in window_values:
                    try:
                        numeric_values.append(float(val))
                    except (ValueError, TypeError):
                        pass
                
                if len(numeric_values) < self.window_size:
                    failed_cells.append({
                        'axis': 1,
                        'column': col_idx,
                        'column_name': df.columns[col_idx],
                        'start_row': start_row,
                        'end_row': end_row - 1,
                        'row_names': window_rows,
                        'reason': '非數字值過多',
                        'values': window_values
                    })
                    continue
                
                # 使用 LLM 偵測規律
                values_matrix = [[v] for v in numeric_values]
                row_names = [str(row) for row in window_rows]
                col_names = [str(df.columns[col_idx])]
                llm_result = self.llm_detector.detect_math_rules(
                    values_matrix,
                    row_names,
                    col_names,
                    use_openai=False
                )
                
                if llm_result and llm_result.get('rules'):
                    # 驗證規律
                    validation = self.llm_detector.validate_rules(
                        values_matrix,
                        llm_result.get('rules')
                    )
                    
                    if validation.get('valid_rules', 0) > 0:
                        valid_windows += 1
                    else:
                        failed_cells.append({
                            'axis': 1,
                            'column': col_idx,
                            'column_name': str(df.columns[col_idx]),
                            'start_row': start_row,
                            'end_row': end_row - 1,
                            'row_names': window_rows,
                            'reason': '規則無法驗證',
                            'values': numeric_values,
                            'rules': [r.get('equation') for r in llm_result.get('rules', [])]
                        })
                else:
                    failed_cells.append({
                        'axis': 1,
                        'column': col_idx,
                        'column_name': str(df.columns[col_idx]),
                        'start_row': start_row,
                        'end_row': end_row - 1,
                        'row_names': window_rows,
                        'reason': '未偵測到規律',
                        'values': numeric_values
                    })
        
        success_rate = valid_windows / total_windows if total_windows > 0 else 0
        
        m_print(f"總窗格數: {total_windows}")
        m_print(f"有效窗格數: {valid_windows}")
        m_print(f"成功率: {success_rate:.2%}")
        m_print(f"失敗位置數: {len(failed_cells)}")
        
        return {
            'success': True,
            'axis': 1,
            'description': llm_analysis.get('axis1_description', ''),
            'confidence': llm_analysis.get('axis1_confidence', 0),
            'total_windows': total_windows,
            'valid_windows': valid_windows,
            'success_rate': success_rate,
            'failed_cells': failed_cells,
            'failed_count': len(failed_cells)
        }
    
    def _load_pattern_result(self, pattern_file: str) -> Optional[Dict[str, Any]]:
        """載入 pattern 結果"""
        pattern_path = Path(pattern_file)
        if not pattern_path.is_absolute():
            base_path = Path(__file__).parent.parent
            pattern_path = base_path / 'prompts' / pattern_file
        
        try:
            with open(pattern_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            m_print(f"無法載入 pattern 結果: {e}")
            return None
    
    def save_result(self, result: Dict[str, Any], output_file: str = None) -> str:
        """保存驗證結果"""
        if not result.get('success'):
            m_print("驗證結果不成功，無法保存")
            return None
        
        if output_file is None:
            prompts_dir = Path(__file__).parent.parent / 'prompts'
            prompts_dir.mkdir(exist_ok=True)
            timestamp = dt.now().strftime('%Y%m%d_%H%M%S')
            output_file = prompts_dir / f'axis_validation_{timestamp}.json'
        else:
            output_file = Path(output_file)
        
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2, default=str)
            
            m_print(f"結果已保存到: {output_file}")
            return str(output_file)
        except Exception as e:
            m_print(f"保存失敗: {e}")
            return None


def main():
    """主函數"""
    parser = argparse.ArgumentParser(description='Axis 方向規律驗證器')
    parser.add_argument('pattern_result', nargs='?', default='height_pattern_result.json',
                       help='pattern 結果檔案（預設: height_pattern_result.json）')
    parser.add_argument('excel_file', nargs='?', default='defect_simple.xlsx',
                       help='Excel檔案路徑')
    parser.add_argument('--window-size', type=int, default=3,
                       help='窗格大小（預設: 3）')
    parser.add_argument('--output', type=str,
                       help='指定輸出檔案路徑')
    
    args = parser.parse_args()
    
    # 初始化驗證器
    validator = AxisPatternValidator(window_size=args.window_size)
    
    # 執行驗證
    result = validator.validate_pattern(args.pattern_result, args.excel_file)
    
    if result.get('success'):
        # 保存結果
        output_file = validator.save_result(result, args.output)
        
        # 顯示摘要
        print(f"\n=== 驗證摘要 ===")
        print(f"Excel檔案: {result.get('excel_file', 'N/A')}")
        print(f"驗證方向: axis={result.get('axis', 'N/A')}")
        print(f"規律描述: {result.get('description', 'N/A')}")
        print(f"LLM信心度: {result.get('confidence', 0):.2%}")
        print(f"\n=== 驗證統計 ===")
        print(f"總窗格數: {result.get('total_windows', 0)}")
        print(f"有效窗格數: {result.get('valid_windows', 0)}")
        print(f"成功率: {result.get('success_rate', 0):.2%}")
        print(f"失敗位置數: {result.get('failed_count', 0)}")
        
        if result.get('failed_count', 0) > 0:
            print(f"\n=== 失敗位置（前10個）===")
            for i, cell in enumerate(result['failed_cells'][:10]):
                print(f"{i+1}. {cell}")
        
        if output_file:
            print(f"\n完整結果已保存到: {output_file}")
    else:
        print(f"驗證失敗: {result.get('error', '未知錯誤')}")


if __name__ == "__main__":
    main()
