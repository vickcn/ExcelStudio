#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
整合偵錯應用腳本
結合數據範圍偵測、規則發現和數據驗證功能
"""

import os
import json
import argparse
import pandas as pd
import numpy as np
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime
from dotenv import load_dotenv
import requests
import openai
import re
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.feature_extraction.text import TfidfVectorizer
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.font_manager import FontProperties

from utils import LOGger

# 載入環境變數
load_dotenv()

class IntegratedDebugApp:
    """整合偵錯應用程式"""
    
    def __init__(self, rules_file: str = None):
        """初始化偵錯應用程式"""
        self.rules_file = rules_file or os.path.join(os.path.dirname(__file__), '..', 'prompts', 'discovered_rules.json')
        self.window_shape = (3, 1)  # 預設窗格大小
        
        # LLM設定
        self.llm_api_url = os.getenv('LLM_API_URL', 'http://localhost:11434/v1').rstrip('/')
        self.llm_model_name = os.getenv('LLM_MODEL_NAME', 'llama3.1')
        self.openai_api_key = os.getenv('OPENAI_API_KEY')
        
        # 初始化客戶端
        self.openai_client = None
        self.session = requests.Session()
        
        if self.openai_api_key and self.openai_api_key != 'your_openai_api_key_here':
            self.openai_client = openai.OpenAI(api_key=self.openai_api_key)
        
        # 設定基本標頭
        base_headers = {
            'Content-Type': 'application/json',
            'User-Agent': 'LLM-API-Client/1.0'
        }
        self.session.headers.update(base_headers)
        
        # 初始化日誌
        self.logger = LOGger.addloger(logfile='')
        
        # 字體設定
        self.font_path = self._find_chinese_font()
        
        # 載入已發現的規則
        self.discovered_rules = self._load_discovered_rules()
        
        # 向量化器（用於索引相似度計算）
        self.vectorizer = TfidfVectorizer()
        
    def _find_chinese_font(self) -> str:
        """尋找中文字體"""
        font_paths = [
            'D:/JobProject/TWSG/CE/重量規格/msjh.ttc',
            'D:/JobProject/TWSG/CE/重量規格/DejaVuSans.ttf',
            'C:/Windows/Fonts/msjh.ttc',
            'C:/Windows/Fonts/simsun.ttc'
        ]
        
        for font_path in font_paths:
            if os.path.exists(font_path):
                return font_path
        
        self.logger("Warning: Chinese font not found, using default font")
        return None
    
    def _load_discovered_rules(self) -> Dict[str, Any]:
        """載入已發現的規則"""
        try:
            with open(self.rules_file, 'r', encoding='utf-8') as f:
                rules = json.load(f)
            self.logger(f"Successfully loaded discovered rules from {self.rules_file}")
            return rules
        except Exception as e:
            self.logger(f"Failed to load discovered rules: {e}")
            return {}
    
    def detect_data_ranges(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """偵測數據範圍"""
        self.logger("開始偵測數據範圍...")
        
        data_ranges = []
        rows, cols = df.shape
        
        # 使用窗格掃描方式偵測數據區域
        window_height, window_width = self.window_shape
        
        for start_row in range(0, rows, window_height):
            for start_col in range(0, cols, window_width):
                end_row = min(start_row + window_height, rows)
                end_col = min(start_col + window_width, cols)
                
                # 檢查是否為數值數據區域
                if self._is_numeric_data_area(df, start_row, start_col, end_row, end_col):
                    # 提取行標題（從第0列）
                    row_names = []
                    if start_col > 0:  # 確保有行標題列
                        for r in range(start_row, end_row):
                            if r < rows:
                                row_name = df.iloc[r, 0] if pd.notna(df.iloc[r, 0]) else f"Row_{r}"
                                row_names.append(str(row_name).strip())
                    
                    data_range = {
                        'start_row': start_row,
                        'end_row': end_row,
                        'start_col': start_col,
                        'end_col': end_col,
                        'row_names': row_names,
                        'window_id': f"window_{start_row}_{start_col}",
                        'data_area': df.iloc[start_row:end_row, start_col:end_col].copy()
                    }
                    data_ranges.append(data_range)
        
        self.logger(f"偵測到 {len(data_ranges)} 個數據範圍")
        return data_ranges
    
    def _is_numeric_data_area(self, df: pd.DataFrame, start_row: int, start_col: int, end_row: int, end_col: int) -> bool:
        """判斷窗格是否包含數值數據區域"""
        numeric_count = 0
        total_non_empty = 0
        
        for i in range(start_row, end_row):
            for j in range(start_col, end_col):
                if i < df.shape[0] and j < df.shape[1]:
                    cell_value = df.iloc[i, j]
                    
                    if pd.notna(cell_value):
                        total_non_empty += 1
                        
                        # 檢查是否為數值
                        if isinstance(cell_value, (int, float)):
                            numeric_count += 1
                        else:
                            # 嘗試轉換為數值
                            try:
                                cell_str = str(cell_value).strip()
                                cleaned = cell_str.replace(',', '').replace('±', '').replace(' ', '')
                                float(cleaned)
                                # 排除包含中文的文字
                                if not any(char.isalpha() and ord(char) > 127 for char in cell_str):
                                    numeric_count += 1
                            except (ValueError, AttributeError):
                                pass
        
        if total_non_empty == 0:
            return False
        
        # 數值比例必須超過70%
        numeric_ratio = numeric_count / total_non_empty
        return numeric_ratio >= 0.7 and numeric_count >= 2
    
    def match_row_indices(self, source_row_names: List[str], target_row_names: List[str]) -> Dict[int, int]:
        """透過相似度計算匹配行索引"""
        self.logger("開始進行索引擬合...")
        
        if not source_row_names or not target_row_names:
            return {}
        
        # 使用TF-IDF向量化
        all_names = source_row_names + target_row_names
        try:
            tfidf_matrix = self.vectorizer.fit_transform(all_names)
            
            source_vectors = tfidf_matrix[:len(source_row_names)]
            target_vectors = tfidf_matrix[len(source_row_names):]
            
            # 計算相似度矩陣
            similarity_matrix = cosine_similarity(target_vectors, source_vectors)
            
            # 找到最佳匹配
            matches = {}
            for target_idx, similarities in enumerate(similarity_matrix):
                best_source_idx = np.argmax(similarities)
                best_similarity = similarities[best_source_idx]
                
                # 只有相似度超過閾值才認為是匹配
                if best_similarity > 0.3:
                    matches[target_idx] = best_source_idx
                    self.logger(f"匹配: {target_row_names[target_idx]} -> {source_row_names[best_source_idx]} (相似度: {best_similarity:.3f})")
            
            return matches
            
        except Exception as e:
            self.logger(f"索引擬合失敗: {e}")
            return {}
    
    def validate_data_with_rules(self, target_df: pd.DataFrame, target_ranges: List[Dict], 
                                source_rules: Dict[str, Any], use_openai: bool = False) -> Dict[str, Any]:
        """使用規則驗證目標數據"""
        self.logger("開始使用規則驗證數據...")
        
        validation_results = {
            'total_validations': 0,
            'passed_validations': 0,
            'failed_validations': 0,
            'anomalies': [],
            'validation_details': []
        }
        
        for range_info in target_ranges:
            row_names = range_info['row_names']
            data_area = range_info['data_area']
            
            # 尋找匹配的規則
            matching_rules = self._find_matching_rules(row_names, source_rules)
            
            if not matching_rules:
                self.logger(f"未找到匹配的規則，跳過窗格 {range_info['window_id']}")
                continue
            
            # 對每個匹配的規則進行驗證
            for rule_key, rule_info in matching_rules.items():
                validation_result = self._validate_with_rule(
                    data_area, row_names, rule_info, range_info, use_openai
                )
                
                validation_results['validation_details'].append(validation_result)
                validation_results['total_validations'] += 1
                
                if validation_result['is_valid']:
                    validation_results['passed_validations'] += 1
                else:
                    validation_results['failed_validations'] += 1
                    # 添加異常位置
                    for anomaly in validation_result.get('anomalies', []):
                        global_anomaly = {
                            'location': (
                                range_info['start_row'] + anomaly['local_row'],
                                range_info['start_col'] + anomaly['local_col']
                            ),
                            'rule_description': rule_info['passed_rules'][0]['description'] if rule_info['passed_rules'] else 'Unknown rule',
                            'window_id': range_info['window_id'],
                            'error_message': anomaly.get('error_message', '')
                        }
                        validation_results['anomalies'].append(global_anomaly)
        
        # 計算總體統計
        if validation_results['total_validations'] > 0:
            success_rate = validation_results['passed_validations'] / validation_results['total_validations']
            self.logger(f"驗證完成: {validation_results['passed_validations']}/{validation_results['total_validations']} 通過 (成功率: {success_rate:.2%})")
        
        return validation_results
    
    def _find_matching_rules(self, row_names: List[str], source_rules: Dict[str, Any]) -> Dict[str, Any]:
        """尋找匹配的規則"""
        matching_rules = {}
        
        for rule_key, rule_info in source_rules.items():
            rule_row_names = rule_info.get('start_loc_row_indicated', [])
            
            if not rule_row_names:
                continue
            
            # 檢查是否有匹配的行名稱
            matches = self.match_row_indices(rule_row_names, row_names)
            
            if matches:
                matching_rules[rule_key] = rule_info
                self.logger(f"找到匹配規則 {rule_key}: {rule_info.get('start_loc_row_name', 'Unknown')}")
        
        return matching_rules
    
    def _validate_with_rule(self, data_area: pd.DataFrame, row_names: List[str], 
                           rule_info: Dict[str, Any], range_info: Dict[str, Any], 
                           use_openai: bool = False) -> Dict[str, Any]:
        """使用特定規則驗證數據"""
        
        validation_result = {
            'rule_key': rule_info.get('start_loc_row_name', 'Unknown'),
            'is_valid': True,
            'anomalies': [],
            'validation_method': 'llm' if use_openai else 'local_llm'
        }
        
        # 準備LLM提示
        prompt = self._prepare_validation_prompt(data_area, row_names, rule_info)
        
        # 呼叫LLM進行驗證
        try:
            if use_openai and self.openai_client:
                response = self._call_openai(prompt)
            else:
                response = self._call_local_llm(prompt)
            
            # 解析LLM回應
            anomalies = self._parse_validation_response(response)
            
            if anomalies:
                validation_result['is_valid'] = False
                validation_result['anomalies'] = anomalies
            
        except Exception as e:
            self.logger(f"LLM驗證失敗: {e}")
            validation_result['is_valid'] = False
            validation_result['error'] = str(e)
        
        return validation_result
    
    def _prepare_validation_prompt(self, data_area: pd.DataFrame, row_names: List[str], 
                                  rule_info: Dict[str, Any]) -> str:
        """準備驗證提示"""
        
        # 獲取規則描述
        rules_text = ""
        if 'passed_rules' in rule_info and rule_info['passed_rules']:
            for i, rule in enumerate(rule_info['passed_rules'], 1):
                rules_text += f"{i}. {rule.get('description', rule.get('rule', 'Unknown rule'))}\n"
                if 'equation' in rule:
                    rules_text += f"   數學公式: {rule['equation']}\n"
        
        # 將數據轉換為易讀格式
        data_str = data_area.to_string(na_rep='空值', max_cols=15)
        
        prompt = f"""
你是一個Excel數據驗證專家。請根據以下已知的數學規則檢查數據表格中的異常：

已知規則（來自歷史數據分析）：
{rules_text}

當前數據表格：
行標題: {row_names}
{data_str}

請仔細檢查當前數據表格中的每個數值，根據行標題和已知規則判斷是否存在異常。

重要注意事項：
- 根據已知規則中的數學關係來驗證數值
- 允許小幅度的數值誤差（±0.1以內）
- 座標使用表格內的相對位置（從0開始）

要求：
1. 回傳格式為包含異常位置和錯誤描述的字典列表
2. 格式：[{{"local_row": 行索引, "local_col": 列索引, "error_message": "錯誤描述"}}, ...]
3. 如果沒有異常，回傳空列表：[]
4. 只回傳JSON格式的列表，不要其他說明文字

範例回傳格式：
[{{"local_row": 0, "local_col": 1, "error_message": "根據規則1，此數值應為XX但實際為YY"}}, {{"local_row": 2, "local_col": 0, "error_message": "違反數學關係式"}}]
"""
        return prompt
    
    def _call_openai(self, prompt: str) -> str:
        """呼叫OpenAI API"""
        try:
            response = self.openai_client.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=1500,
                timeout=120
            )
            return response.choices[0].message.content
        except Exception as e:
            self.logger(f"OpenAI call failed: {e}")
            return "[]"
    
    def _call_local_llm(self, prompt: str) -> str:
        """呼叫本地LLM"""
        try:
            payload = {
                'model': self.llm_model_name,
                'messages': [{"role": "user", "content": prompt}],
                'temperature': 0.1,
                'max_tokens': 1000
            }
            
            headers = {'Content-Type': 'application/json'}
            api_key = os.getenv('LLM_API_KEY')
            if api_key:
                headers['Authorization'] = f'Bearer {api_key}'
            
            response = self.session.post(
                f"{self.llm_api_url}/chat/completions",
                json=payload,
                headers=headers,
                timeout=120
            )
            
            if response.status_code == 200:
                result = response.json()
                if 'choices' in result and len(result['choices']) > 0:
                    return result['choices'][0]['message']['content']
                elif 'response' in result:
                    return result['response']
            
            self.logger(f"Local LLM response error: {response.text}")
            return "[]"
            
        except Exception as e:
            self.logger(f"Local LLM call failed: {e}")
            return "[]"
    
    def _parse_validation_response(self, response: str) -> List[Dict[str, Any]]:
        """解析驗證回應"""
        try:
            # 清理回應文字
            response = response.strip()
            
            # 尋找JSON格式
            start_idx = response.find('[')
            end_idx = response.rfind(']') + 1
            
            if start_idx != -1 and end_idx > start_idx:
                json_str = response[start_idx:end_idx]
                anomalies = json.loads(json_str)
                
                # 驗證格式
                result = []
                for item in anomalies:
                    if isinstance(item, dict) and 'local_row' in item and 'local_col' in item:
                        result.append({
                            'local_row': int(item['local_row']),
                            'local_col': int(item['local_col']),
                            'error_message': item.get('error_message', 'Unknown error')
                        })
                return result
            
            return []
            
        except Exception as e:
            self.logger(f"Failed to parse validation response: {e}")
            return []
    
    def visualize_results(self, df: pd.DataFrame, anomalies: List[Dict], output_path: str):
        """視覺化驗證結果"""
        try:
            # 設定matplotlib中文字體
            if self.font_path:
                font_prop = FontProperties(fname=self.font_path)
                plt.rcParams['font.family'] = font_prop.get_name()
            
            # 創建圖形
            fig, ax = plt.subplots(figsize=(max(12, df.shape[1] * 1.5), max(8, df.shape[0] * 0.8)))
            
            # 隱藏軸
            ax.set_xlim(0, df.shape[1])
            ax.set_ylim(0, df.shape[0])
            ax.invert_yaxis()
            
            # 創建異常位置集合
            anomaly_locations = {anomaly['location'] for anomaly in anomalies}
            
            # 繪製表格
            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    # 判斷是否為異常位置
                    is_anomaly = (i, j) in anomaly_locations
                    
                    # 設定顏色
                    if is_anomaly:
                        bg_color = 'red'
                        text_color = 'white'
                    else:
                        bg_color = 'white'
                        text_color = 'black'
                    
                    # 繪製儲存格背景
                    rect = patches.Rectangle((j, i), 1, 1, linewidth=1, 
                                           edgecolor='black', facecolor=bg_color)
                    ax.add_patch(rect)
                    
                    # 添加儲存格內容
                    cell_value = df.iloc[i, j]
                    if pd.notna(cell_value):
                        text = str(cell_value)
                        if len(text) > 15:
                            text = text[:15] + '...'
                        
                        ax.text(j + 0.5, i + 0.5, text, 
                               ha='center', va='center', 
                               color=text_color, fontsize=8,
                               fontproperties=font_prop if self.font_path else None)
            
            # 設定標題
            title = f'數據驗證結果\n(發現 {len(anomalies)} 個異常)'
            plt.title(title, fontproperties=font_prop if self.font_path else None, fontsize=12)
            
            # 移除軸標籤
            ax.set_xticks([])
            ax.set_yticks([])
            
            # 儲存圖片
            plt.tight_layout()
            plt.savefig(output_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            self.logger(f"視覺化結果已儲存至: {output_path}")
            
        except Exception as e:
            self.logger(f"視覺化失敗: {e}")
    
    def run_debug_analysis(self, target_file: str, use_openai: bool = False, 
                          output_dir: str = './debug_results') -> Dict[str, Any]:
        """執行完整的偵錯分析"""
        self.logger(f"開始偵錯分析: {target_file}")
        
        # 創建輸出目錄
        os.makedirs(output_dir, exist_ok=True)
        
        try:
            # 1. 載入目標數據
            target_df = pd.read_excel(target_file)
            self.logger(f"載入目標文件，形狀: {target_df.shape}")
            
            # 2. 偵測數據範圍
            target_ranges = self.detect_data_ranges(target_df)
            
            if not target_ranges:
                self.logger("未偵測到有效的數據範圍")
                return {'error': '未偵測到有效的數據範圍'}
            
            # 3. 使用規則驗證數據
            validation_results = self.validate_data_with_rules(
                target_df, target_ranges, self.discovered_rules, use_openai
            )
            
            # 4. 準備完整結果
            results = {
                'target_file': target_file,
                'analysis_time': datetime.now().isoformat(),
                'data_ranges_found': len(target_ranges),
                'validation_results': validation_results,
                'summary': {
                    'total_validations': validation_results['total_validations'],
                    'passed_validations': validation_results['passed_validations'],
                    'failed_validations': validation_results['failed_validations'],
                    'anomalies_count': len(validation_results['anomalies']),
                    'success_rate': validation_results['passed_validations'] / max(validation_results['total_validations'], 1)
                }
            }
            
            # 5. 儲存結果
            result_file = os.path.join(output_dir, 'debug_analysis_result.json')
            with open(result_file, 'w', encoding='utf-8') as f:
                # 轉換tuple為list以便JSON序列化
                json_results = self._prepare_json_results(results)
                json.dump(json_results, f, ensure_ascii=False, indent=2)
            
            self.logger(f"分析結果已儲存至: {result_file}")
            
            # 6. 視覺化結果
            if validation_results['anomalies']:
                output_image = os.path.join(output_dir, 'debug_visualization.png')
                self.visualize_results(target_df, validation_results['anomalies'], output_image)
            
            return results
            
        except Exception as e:
            self.logger(f"偵錯分析失敗: {e}")
            return {'error': str(e)}
    
    def _prepare_json_results(self, results: Dict[str, Any]) -> Dict[str, Any]:
        """準備JSON序列化的結果"""
        json_results = results.copy()
        
        # 轉換tuple為list
        if 'validation_results' in json_results and 'anomalies' in json_results['validation_results']:
            for anomaly in json_results['validation_results']['anomalies']:
                if 'location' in anomaly and isinstance(anomaly['location'], tuple):
                    anomaly['location'] = list(anomaly['location'])
        
        return json_results

def main():
    """主程式"""
    parser = argparse.ArgumentParser(description='整合偵錯應用程式')
    parser.add_argument('target_file', help='要驗證的目標Excel文件路徑')
    parser.add_argument('--rules-file', help='規則文件路徑（預設使用discovered_rules.json）')
    parser.add_argument('--use-openai', action='store_true', help='使用OpenAI而非本地LLM')
    parser.add_argument('--output-dir', default=r'./tmp2/dataProcess/debug_results', help='輸出目錄')
    parser.add_argument('--window-shape', nargs=2, type=int, default=[5, 1], 
                       help='窗格大小 (行數 列數)，預設為 5 1')
    
    args = parser.parse_args()
    
    # 檢查文件是否存在
    if not os.path.exists(args.target_file):
        print(f"錯誤: 找不到文件 - {args.target_file}")
        return
    
    try:
        # 初始化偵錯應用程式
        debug_app = IntegratedDebugApp(args.rules_file)
        debug_app.window_shape = tuple(args.window_shape)
        
        # 執行偵錯分析
        results = debug_app.run_debug_analysis(
            args.target_file, 
            args.use_openai, 
            args.output_dir
        )
        
        # 輸出結果摘要
        if 'error' in results:
            print(f"分析失敗: {results['error']}")
        else:
            summary = results['summary']
            print(f"\n=== 偵錯分析結果 ===")
            print(f"目標文件: {results['target_file']}")
            print(f"偵測到數據範圍: {results['data_ranges_found']}")
            print(f"總驗證次數: {summary['total_validations']}")
            print(f"通過驗證: {summary['passed_validations']}")
            print(f"失敗驗證: {summary['failed_validations']}")
            print(f"發現異常: {summary['anomalies_count']}")
            print(f"成功率: {summary['success_rate']:.2%}")
            
            if summary['anomalies_count'] > 0:
                print(f"\n異常位置:")
                for i, anomaly in enumerate(results['validation_results']['anomalies'], 1):
                    location = anomaly['location']
                    rule_desc = anomaly['rule_description']
                    print(f"  {i}. 位置 {location}: {rule_desc}")
        
    except Exception as e:
        print(f"執行失敗: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()

