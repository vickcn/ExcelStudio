#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
高度規律偵測器 - 偵測Excel表格中數字的大小規律
"""

import os
import json
import pandas as pd
import numpy as np
from typing import List, Dict, Any, Optional
from pathlib import Path
import argparse
from datetime import datetime as dt
import requests
from dotenv import load_dotenv

from utils import LOGger

m_logfile = Path(__file__).resolve().parent.parent / "outputs" / "detect_height_pattern.log"
m_print = LOGger.addloger(logfile=str(m_logfile))

# 載入環境變數
load_dotenv()

# 導入自定義模組
from universal_table_detector import UniversalTableDetector


class HeightPatternDetector:
    """高度規律偵測器"""
    
    def __init__(self, provider: str = None, model: str = None):
        """
        初始化偵測器
        
        Args:
            provider: TextProcessor provider (如果為None則從config.json讀取)
            model: TextProcessor model alias (如果為None則從config.json讀取)
        """
        # 從 config.json 讀取設定
        config_path = Path(__file__).parent.parent.parent / 'config.json'
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    textprocessor_config = config.get('textprocessor', {})
                    self.textprocessor_url = textprocessor_config.get('url', 'http://localhost:6017/chat')
                    self.provider = provider or textprocessor_config.get('provider', 'remote')
                    self.model = model or textprocessor_config.get('model', 'remote8b')
                    self.timeout = textprocessor_config.get('timeout', 120)
                    self.max_tokens = textprocessor_config.get('max_tokens', 2000)
                    self.temperature = textprocessor_config.get('temperature', 0.1)
                    m_print(f"已從 config.json 載入設定")
            except Exception as e:
                m_print(f"無法載入 config.json: {e}，使用預設值")
                self.textprocessor_url = 'http://localhost:6017/chat'
                self.provider = provider or 'remote'
                self.model = model or 'remote8b'
                self.timeout = 120
                self.max_tokens = 2000
                self.temperature = 0.1
        else:
            m_print(f"找不到 config.json，使用預設值")
            self.textprocessor_url = 'http://localhost:6017/chat'
            self.provider = provider or 'remote'
            self.model = model or 'remote8b'
            self.timeout = 120
            self.max_tokens = 2000
            self.temperature = 0.1
        
        # 載入 prompt 模板
        self.prompt_template = self._load_prompt_template()
        
        # 初始化表格偵測器
        self.table_detector = UniversalTableDetector()
        
        m_print(f"高度規律偵測器已初始化")
        m_print(f"TextProcessor URL: {self.textprocessor_url}")
        m_print(f"Provider: {self.provider}, Model: {self.model}")
    
    def _load_prompt_template(self) -> Dict[str, Any]:
        """載入 prompt 模板"""
        prompt_file = Path(__file__).parent.parent / 'prompts' / 'height_pattern_detection.json'
        
        try:
            with open(prompt_file, 'r', encoding='utf-8') as f:
                template = json.load(f)
            m_print(f"已載入 prompt 模板: {prompt_file}")
            return template
        except Exception as e:
            m_print(f"無法載入 prompt 模板: {e}")
            return {
                "system_prompt": "分析表格數據",
                "user_prompt_template": "",
                "description": "默認模板"
            }
    
    def analyze_excel_file(self, excel_file: str) -> Dict[str, Any]:
        """
        分析Excel檔案中的高度規律
        
        Args:
            excel_file: Excel檔案路徑
            
        Returns:
            Dict: 分析結果
        """
        excel_path = Path(excel_file)
        if not excel_path.is_absolute():
            base_path = Path(__file__).parent.parent.parent
            excel_path = base_path / excel_file
        
        if not excel_path.exists():
            m_print(f"找不到Excel檔案: {excel_path}")
            return {'success': False, 'error': f'找不到檔案: {excel_path}'}
        
        m_print(f"開始分析Excel檔案: {excel_path}")
        
        try:
            # 步驟1: 偵測表格結構
            m_print("=== 步驟1: 偵測表格結構 ===")
            table_results = self.table_detector.detect_tables_by_analysis(
                str(excel_path), 
                table_mode='pure_numeric',
                use_llm=False
            )
            
            if not table_results:
                m_print("未偵測到有效表格")
                return {'success': False, 'error': '未偵測到有效表格'}
            
            m_print(f"偵測到 {len(table_results)} 個表格")
            
            # 步驟2: 提取DataFrame
            m_print("=== 步驟2: 提取DataFrame ===")
            dataframes = self.table_detector.extract_dataframes(str(excel_path), table_results)
            
            if not dataframes:
                m_print("無法提取DataFrame")
                return {'success': False, 'error': '無法提取DataFrame'}
            
            # 使用第一個DataFrame進行分析
            main_df = dataframes[0]
            m_print(f"使用DataFrame形狀: {main_df.shape}")
            
            # 步驟3: 準備axis數據
            m_print("=== 步驟3: 準備axis數據 ===")
            
            # axis=0: 沿行的方向（每行）
            axis0_data = []
            skipped_cells_axis0 = 0
            for idx, row in main_df.iterrows():
                row_values = []
                for x in row.values:
                    if pd.notna(x) and str(x).strip() != '':
                        try:
                            row_values.append(float(x))
                        except (ValueError, TypeError):
                            skipped_cells_axis0 += 1
                            continue
                
                if row_values:
                    axis0_data.append({
                        'row_index': str(idx),
                        'values': row_values,
                        'min': min(row_values),
                        'max': max(row_values),
                        'range': max(row_values) - min(row_values),
                        'mean': np.mean(row_values)
                    })
            
            # axis=1: 沿列的方向（每列）
            axis1_data = []
            skipped_cells_axis1 = 0
            for col in main_df.columns:
                col_values = []
                for x in main_df[col].values:
                    if pd.notna(x) and str(x).strip() != '':
                        try:
                            col_values.append(float(x))
                        except (ValueError, TypeError):
                            skipped_cells_axis1 += 1
                            continue
                
                if col_values:
                    axis1_data.append({
                        'column': str(col),
                        'values': col_values,
                        'min': min(col_values),
                        'max': max(col_values),
                        'range': max(col_values) - min(col_values),
                        'mean': np.mean(col_values)
                    })
            
            if skipped_cells_axis0 > 0:
                m_print(f"axis=0 跳過了 {skipped_cells_axis0} 個非數字儲存格")
            if skipped_cells_axis1 > 0:
                m_print(f"axis=1 跳過了 {skipped_cells_axis1} 個非數字儲存格")
            
            m_print(f"準備了 {len(axis0_data)} 行數據（axis=0）")
            m_print(f"準備了 {len(axis1_data)} 列數據（axis=1）")
            
            # 步驟4: 使用LLM分析
            m_print("=== 步驟4: LLM分析高度規律 ===")
            analysis_result = self._analyze_with_llm(axis0_data, axis1_data)
            
            if not analysis_result:
                m_print("LLM分析失敗")
                return {'success': False, 'error': 'LLM分析失敗'}
            
            # 組裝完整結果
            result = {
                'success': True,
                'excel_file': str(excel_path),
                'analysis_time': dt.now().isoformat(),
                'table_detection': {
                    'tables_found': len(table_results),
                    'dataframe_shape': main_df.shape,
                    'rows': len(axis0_data),
                    'columns': len(axis1_data)
                },
                'axis0_data': axis0_data[:5],  # 只保存前5行
                'axis1_data': axis1_data[:5],  # 只保存前5列
                'llm_analysis': analysis_result
            }
            
            m_print(f"\n=== 分析完成 ===")
            m_print(f"axis=0 規律: {analysis_result.get('axis0_has_pattern', 'N/A')}")
            m_print(f"axis=0 信心值: {analysis_result.get('axis0_confidence', 0):.2%}")
            m_print(f"axis=1 規律: {analysis_result.get('axis1_has_pattern', 'N/A')}")
            m_print(f"axis=1 信心值: {analysis_result.get('axis1_confidence', 0):.2%}")
            
            return result
            
        except Exception as e:
            m_print(f"分析過程發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            return {'success': False, 'error': str(e)}
    
    def _analyze_with_llm(self, axis0_data: List[Dict], axis1_data: List[Dict]) -> Optional[Dict[str, Any]]:
        """
        使用LLM分析高度規律
        
        Args:
            axis0_data: 行數據
            axis1_data: 列數據
            
        Returns:
            Dict: LLM分析結果
        """
        try:
            # 準備數據字符串
            axis0_str = self._format_data_for_prompt(axis0_data)
            axis1_str = self._format_data_for_prompt(axis1_data)
            
            # 構建提示
            system_prompt = self.prompt_template.get('system_prompt', '')
            user_prompt_template = self.prompt_template.get('user_prompt_template', '')
            
            user_prompt = user_prompt_template.format(
                axis0_data=axis0_str,
                axis1_data=axis1_str
            )
            
            m_print(f"發送LLM請求到 TextProcessor...")
            
            # 調用TextProcessor的 /chat 端點
            payload = {
                'prompt': user_prompt,
                'provider': self.provider,
                'model': self.model,
                'system_prompt': system_prompt,
                'max_tokens': self.max_tokens,
                'temperature': self.temperature
            }
            
            response = requests.post(
                self.textprocessor_url,
                json=payload,
                timeout=self.timeout
            )
            
            if response.status_code != 200:
                m_print(f"TextProcessor API錯誤: {response.status_code} - {response.text}")
                return None
            
            result = response.json()
            
            # 檢查錯誤
            if 'error' in result:
                m_print(f"TextProcessor回傳錯誤: {result['error']}")
                return None
            
            # 提取內容
            content = result.get('output', '')
            if not content:
                m_print("TextProcessor回應為空")
                return None
            
            m_print(f"LLM回應: {content[:200]}...")
            
            # 解析JSON
            try:
                # 嘗試提取JSON部分
                json_match = self._extract_json_from_response(content)
                if json_match:
                    analysis_result = json.loads(json_match)
                    m_print(f"成功解析LLM回應")
                    return analysis_result
                else:
                    m_print(f"無法從回應中提取JSON")
                    m_print(f"原始回應: {content}")
                    return None
            except json.JSONDecodeError as e:
                m_print(f"JSON解析失敗: {e}")
                m_print(f"原始內容: {content}")
                return None
                
        except Exception as e:
            m_print(f"LLM分析發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _format_data_for_prompt(self, data: List[Dict]) -> str:
        """
        格式化數據供LLM使用
        
        Args:
            data: 數據列表
            
        Returns:
            str: 格式化後的字符串
        """
        lines = []
        for i, item in enumerate(data[:10]):  # 只取前10個
            idx = item.get('row_index', item.get('column', str(i)))
            values = item.get('values', [])
            if len(values) > 5:
                values_str = f"{values[:3]} ... {values[-2:]} ({len(values)} 項)"
            else:
                values_str = str(values)
            
            stats = f"min={item.get('min', 0):.2f}, max={item.get('max', 0):.2f}, 範圍={item.get('range', 0):.2f}"
            lines.append(f"{idx}: {values_str} ({stats})")
        
        return "\n".join(lines)
    
    def _extract_json_from_response(self, content: str) -> Optional[str]:
        """
        從LLM回應中提取JSON部分
        
        Args:
            content: LLM回應字符串
            
        Returns:
            Optional[str]: JSON字符串
        """
        import re
        
        # 嘗試找到JSON對象
        json_match = re.search(r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', content)
        if json_match:
            return json_match.group(0)
        
        # 如果找不到，嘗試尋找多行JSON
        json_match = re.search(r'\{[\s\S]*\}', content)
        if json_match:
            return json_match.group(0)
        
        return None
    
    def save_result(self, result: Dict[str, Any], output_file: str = None) -> str:
        """
        保存分析結果
        
        Args:
            result: 分析結果
            output_file: 輸出檔案路徑
            
        Returns:
            str: 保存的檔案路徑
        """
        if not result.get('success'):
            m_print("分析結果不成功，無法保存")
            return None
        
        if output_file is None:
            prompts_dir = Path(__file__).parent.parent / 'prompts'
            prompts_dir.mkdir(exist_ok=True)
            # timestamp = dt.now().strftime('%Y%m%d_%H%M%S')
            output_file = prompts_dir / f'height_pattern_result.json'
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
    parser = argparse.ArgumentParser(description='高度規律偵測器')
    parser.add_argument('excel_file', nargs='?', default='correct_simple.xlsx',
                       help='Excel檔案路徑（預設: correct_simple.xlsx）')
    parser.add_argument('--provider', type=str, default='remote',
                       help='TextProcessor provider (預設: remote)')
    parser.add_argument('--model', type=str, default='remote8b',
                       help='TextProcessor model (預設: remote8b)')
    parser.add_argument('--output', type=str,
                       help='指定輸出檔案路徑')
    
    args = parser.parse_args()
    
    # 初始化偵測器
    detector = HeightPatternDetector(provider=args.provider, model=args.model)
    
    # 執行分析
    result = detector.analyze_excel_file(args.excel_file)
    
    if result['success']:
        # 保存結果
        output_file = detector.save_result(result, args.output)
        
        # 顯示摘要
        print(f"\n=== 分析摘要 ===")
        print(f"檔案: {result['excel_file']}")
        print(f"分析時間: {result['analysis_time']}")
        print(f"表格大小: {result['table_detection']['dataframe_shape']}")
        print(f"\n=== 高度規律偵測結果 ===")
        
        llm_analysis = result['llm_analysis']
        if llm_analysis.get('axis0_has_pattern'):
            print(f"axis=0（行方向）: 有規律")
            print(f"  信心值: {llm_analysis.get('axis0_confidence', 0):.2%}")
            print(f"  描述: {llm_analysis.get('axis0_description', 'N/A')}")
        else:
            print(f"axis=0（行方向）: 無明顯規律")
            print(f"  信心值: {llm_analysis.get('axis0_confidence', 0):.2%}")
        
        if llm_analysis.get('axis1_has_pattern'):
            print(f"axis=1（列方向）: 有規律")
            print(f"  信心值: {llm_analysis.get('axis1_confidence', 0):.2%}")
            print(f"  描述: {llm_analysis.get('axis1_description', 'N/A')}")
        else:
            print(f"axis=1（列方向）: 無明顯規律")
            print(f"  信心值: {llm_analysis.get('axis1_confidence', 0):.2%}")
        
        if output_file:
            print(f"\n詳細結果已保存到: {output_file}")
    else:
        print(f"分析失敗: {result.get('error', '未知錯誤')}")


if __name__ == "__main__":
    main()
