#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel數據異常偵測器
根據valid_rules.json中的規則，偵測Excel表格中的異常數據
"""

import argparse
import pandas as pd
import numpy as np
import json
import os
import requests
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.font_manager import FontProperties
from dotenv import load_dotenv
import re
from typing import List, Tuple, Dict, Any

from utils import LOGger

# 設定日誌
m_print = LOGger.addloger(logfile='')

# 載入環境變數
load_dotenv()

class ExcelAnomalyDetector:
    def __init__(self, rules_file: str = None):
        """初始化異常偵測器"""
        self.rules_file = rules_file or os.path.join(os.path.dirname(__file__), '..', 'prompts', 'valid_rules.json')
        self.rules = self._load_rules()
        self.window_shape = (5, 1)  # 預設窗格大小
        
        # LLM設定
        self.llm_api_url = os.getenv('LLM_API_URL', 'http://localhost:11434/v1').rstrip('/')
        self.llm_model_name = os.getenv('LLM_MODEL_NAME', 'llama3.1')
        self.openai_api_key = os.getenv('OPENAI_API_KEY')
        
        # 字體設定
        self.font_path = self._find_chinese_font()
        
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
        
        m_print("Warning: Chinese font not found, using default font", stamps=['font'])
        return None
        
    def _load_rules(self) -> List[str]:
        """載入驗證規則"""
        try:
            with open(self.rules_file, 'r', encoding='utf-8') as f:
                rules = json.load(f)
            m_print(f"Successfully loaded {len(rules)} rules", stamps=['rules'])
            return rules
        except Exception as e:
            m_print(f"Failed to load rules: {e}", stamps=['rules', 'error'])
            return []
    
    def load_excel_data(self, file_path: str) -> pd.DataFrame:
        """載入Excel數據"""
        try:
            # 嘗試不同的讀取方式
            try:
                df = pd.read_excel(file_path, engine='openpyxl')
            except:
                df = pd.read_excel(file_path, engine='xlrd')
            
            m_print(f"Successfully loaded Excel file, shape: {df.shape}", stamps=['excel'])
            return df
        except Exception as e:
            m_print(f"Failed to load Excel file: {e}", stamps=['excel', 'error'])
            raise
    
    def _is_numeric_data_area(self, df: pd.DataFrame, start_row: int, start_col: int, end_row: int, end_col: int) -> bool:
        """判斷窗格是否包含數值數據區域（更嚴格的條件）"""
        numeric_count = 0
        text_count = 0
        total_non_empty = 0
        
        for i in range(start_row, end_row):
            for j in range(start_col, end_col):
                if i < df.shape[0] and j < df.shape[1]:
                    cell_value = df.iloc[i, j]
                    
                    if pd.notna(cell_value):
                        total_non_empty += 1
                        cell_str = str(cell_value).strip()
                        
                        # 檢查是否為純數值
                        if isinstance(cell_value, (int, float)):
                            numeric_count += 1
                        else:
                            # 嘗試轉換為數值（更嚴格的檢查）
                            try:
                                # 移除常見的數值格式符號
                                cleaned = cell_str.replace(',', '').replace('±', '').replace(' ', '')
                                float(cleaned)
                                # 確保不是純文字（如產品名稱、材質等）
                                if not any(char.isalpha() and ord(char) > 127 for char in cell_str):  # 排除中文字符
                                    numeric_count += 1
                                else:
                                    text_count += 1
                            except (ValueError, AttributeError):
                                text_count += 1
                                # 檢查是否包含明顯的文字標識符
                                text_indicators = ['材質', '砂芯', '重', '後', '量', '客供', '計算', '加', '電鍍', '磨光', '研磨', 'SS', 'RH', 'SW', 'LW']
                                if any(indicator in cell_str for indicator in text_indicators):
                                    text_count += 2  # 文字區域權重加倍
        
        if total_non_empty == 0:
            return False
            
        # 更嚴格的條件：
        # 1. 數值比例必須超過70%
        # 2. 文字數量不能太多
        # 3. 必須至少有2個數值
        numeric_ratio = numeric_count / total_non_empty
        has_enough_numbers = numeric_count >= 2
        not_too_much_text = text_count <= total_non_empty * 0.3
        
        is_data_area = numeric_ratio >= 0.7 and has_enough_numbers and not_too_much_text
        
        # 調試信息
        if total_non_empty > 0:
            m_print(f"Window ({start_row},{start_col}): numeric={numeric_count}, text={text_count}, ratio={numeric_ratio:.2f}, is_data={is_data_area}", stamps=['data_filter'])
        
        return is_data_area
    
    def create_windows(self, df: pd.DataFrame, window_shape: Tuple[int, int] = None) -> List[Dict]:
        """將DataFrame分割成窗格，優先處理數值數據區域"""
        if window_shape is None:
            window_shape = self.window_shape
        
        rows, cols = window_shape
        windows = []
        
        df_rows, df_cols = df.shape
        
        for start_row in range(0, df_rows, rows):
            for start_col in range(0, df_cols, cols):
                end_row = min(start_row + rows, df_rows)
                end_col = min(start_col + cols, df_cols)
                
                # 檢查是否為數值數據區域
                is_data_area = self._is_numeric_data_area(df, start_row, start_col, end_row, end_col)
                
                # 提取子DataFrame
                sub_df = df.iloc[start_row:end_row, start_col:end_col].copy()
                
                # 設定有意義的行標題：使用第0列（Unnamed: 0）的內容作為行標題
                if 'Unnamed: 0' in df.columns and start_col > 0:
                    # 獲取對應行的標題文字
                    row_labels = df.iloc[start_row:end_row, 0].values
                    # 清理標題文字，移除換行符等
                    cleaned_labels = [str(label).replace('\n', ' ').strip() if pd.notna(label) else f"Row_{i}" 
                                    for i, label in enumerate(row_labels)]
                    sub_df.index = cleaned_labels
                
                # 保留原始的列標題
                # columns已經自動保留了原始的列名
                
                window_info = {
                    'sub_df': sub_df,
                    'start_row': start_row,
                    'start_col': start_col,
                    'end_row': end_row,
                    'end_col': end_col,
                    'window_id': f"window_{start_row}_{start_col}",
                    'is_data_area': is_data_area,
                    'priority': 1 if is_data_area else 0  # 數據區域優先處理
                }
                
                windows.append(window_info)
        
        # 按優先級排序，數據區域優先
        windows.sort(key=lambda x: x['priority'], reverse=True)
        
        data_windows = [w for w in windows if w['is_data_area']]
        text_windows = [w for w in windows if not w['is_data_area']]
        
        m_print(f"Created {len(windows)} windows ({len(data_windows)} data areas, {len(text_windows)} text areas)", stamps=['windows'])
        return windows
    
    def prepare_llm_prompt(self, sub_df: pd.DataFrame, window_info: Dict) -> str:
        """準備LLM提示，使用已設定有意義index的子DataFrame"""
        
        # 子DataFrame已經設定了有意義的行標題（來自第0列的內容）
        # 將DataFrame轉換為易讀格式，現在index就是有意義的標題
        sub_df_str = sub_df.to_string(na_rep='空值', max_cols=15)
        
        rules_str = '\n'.join([f"{i+1}. {rule}" for i, rule in enumerate(self.rules)])
        
        prompt = f"""
你是一個Excel數據驗證專家。請根據以下規則檢查數據表格中的異常：

驗證規則：
{rules_str}

當前分析窗格（窗格ID: {window_info['window_id']}）：
{sub_df_str}

請仔細檢查當前分析窗格中的每個數值，根據行標題（index）和列標題（columns）的信息來理解數值含義，並判斷是否違反上述規則。

行標題說明：
- 行標題已經顯示了數值的含義（如"磨光後重"、"研磨量"、"電鍍後重量"等）
- 列標題顯示了不同的產品型號或規格
- 請根據這些標題來理解每個數值代表什麼

重要注意事項：
- 只檢查當前窗格內包含數值的儲存格
- 根據行標題理解數值含義，應用相應的驗證規則
- 根據規則中的重量關係公式來判斷數值是否合理
- 座標使用窗格內的相對位置（從0開始，對應窗格DataFrame的行列索引）

要求：
1. 回傳格式為包含位置和違反規則的字典列表
2. 格式：[{{"location": [行索引, 列索引], "violated_rule": 規則編號, "rule_description": "規則描述"}}, ...]
3. 行索引和列索引是窗格內的相對位置（從0開始）
4. 規則編號對應上述編號（1-{len(self.rules)}）
5. 如果沒有異常，回傳空列表：[]
6. 只回傳JSON格式的列表，不要其他說明文字

範例回傳格式：
[{{"location": [0, 0], "violated_rule": 1, "rule_description": "磨光後重 - 研磨量 = 電鍍後重量"}}, {{"location": [2, 1], "violated_rule": 3, "rule_description": "電鍍後重量 - 磨光後重 = 電鍍重量"}}]
"""
        return prompt
    
    def call_local_llm(self, prompt: str) -> str:
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
            
            response = requests.post(
                self.llm_api_url,
                json=payload,
                headers=headers,
                timeout=1200
            )
            
            if response.status_code == 200:
                result = response.json()
                if 'choices' in result and len(result['choices']) > 0:
                    return result['choices'][0]['message']['content']
                elif 'response' in result:
                    return result['response']
            
            m_print(f"Local LLM response error: {response.text}", stamps=['llm', 'error'])
            return "[]"
            
        except Exception as e:
            m_print(f"Local LLM call failed: {e}", stamps=['llm', 'error'])
            return "[]"
    
    def call_openai(self, prompt: str) -> str:
        """呼叫OpenAI API"""
        try:
            import openai
            client = openai.OpenAI(
                api_key=self.openai_api_key,
                timeout=1200
            )
            
            response = client.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=1500,  # 增加token上限
                timeout=1200
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            m_print(f"OpenAI call failed: {e}", stamps=['openai', 'error'])
            return "[]"
    
    def _validate_anomaly_location(self, df: pd.DataFrame, location: Tuple[int, int], window_info: Dict) -> bool:
        """驗證異常位置是否為有效的數值數據"""
        row, col = location
        global_row = window_info['start_row'] + row
        global_col = window_info['start_col'] + col
        
        # 檢查座標是否在範圍內
        if global_row >= df.shape[0] or global_col >= df.shape[1]:
            return False
            
        cell_value = df.iloc[global_row, global_col]
        
        # 必須是非空值
        if pd.isna(cell_value):
            return False
            
        cell_str = str(cell_value).strip()
        
        # 檢查是否為數值
        if isinstance(cell_value, (int, float)):
            return True
            
        # 嘗試轉換為數值
        try:
            cleaned = cell_str.replace(',', '').replace('±', '').replace(' ', '')
            float(cleaned)
            # 排除包含中文的文字
            if any(char.isalpha() and ord(char) > 127 for char in cell_str):
                return False
            return True
        except (ValueError, AttributeError):
            return False
    
    def parse_llm_response(self, response: str, df: pd.DataFrame = None, window_info: Dict = None) -> List[Dict]:
        """解析LLM回應，返回包含位置和違反規則的詳細信息，並進行二次驗證"""
        try:
            # 清理回應文字
            response = response.strip()
            
            # 嘗試直接解析JSON
            try:
                import json
                # 尋找JSON格式
                start_idx = response.find('[')
                end_idx = response.rfind(']') + 1
                
                if start_idx != -1 and end_idx > start_idx:
                    json_str = response[start_idx:end_idx]
                    anomalies = json.loads(json_str)
                    
                    # 驗證格式並標準化
                    result = []
                    for item in anomalies:
                        if isinstance(item, dict) and 'location' in item:
                            location = item['location']
                            if isinstance(location, list) and len(location) == 2:
                                location_tuple = (int(location[0]), int(location[1]))
                                
                                # 二次驗證：檢查是否為有效的數值位置
                                if df is not None and window_info is not None:
                                    if not self._validate_anomaly_location(df, location_tuple, window_info):
                                        m_print(f"Filtered out non-numeric anomaly at {location_tuple}", stamps=['filter'])
                                        continue
                                
                                result.append({
                                    'location': location_tuple,
                                    'violated_rule': item.get('violated_rule', 0),
                                    'rule_description': item.get('rule_description', 'Unknown rule')
                                })
                    return result
            except json.JSONDecodeError:
                pass
            
            # 回退到舊格式解析（只有座標）
            import re
            pattern = r'\[(.*?)\]'
            matches = re.findall(pattern, response, re.DOTALL)
            
            if not matches:
                return []
            
            # 解析座標
            coords_str = matches[0]
            if not coords_str.strip():
                return []
            
            # 使用eval安全解析（在受控環境下）
            try:
                coords = eval(f"[{coords_str}]")
                if isinstance(coords, list):
                    result = []
                    for coord in coords:
                        if isinstance(coord, (list, tuple)) and len(coord) == 2:
                            location_tuple = (int(coord[0]), int(coord[1]))
                            
                            # 二次驗證
                            if df is not None and window_info is not None:
                                if not self._validate_anomaly_location(df, location_tuple, window_info):
                                    continue
                            
                            result.append({
                                'location': location_tuple,
                                'violated_rule': 0,
                                'rule_description': 'Rule not specified (legacy format)'
                            })
                    return result
            except:
                pass
            
            # 手動解析舊格式
            result = []
            parts = coords_str.split('),')
            for part in parts:
                part = part.strip().replace('(', '').replace(')', '')
                if ',' in part:
                    try:
                        r, c = part.split(',')
                        location_tuple = (int(r.strip()), int(c.strip()))
                        
                        # 二次驗證
                        if df is not None and window_info is not None:
                            if not self._validate_anomaly_location(df, location_tuple, window_info):
                                continue
                        
                        result.append({
                            'location': location_tuple,
                            'violated_rule': 0,
                            'rule_description': 'Rule not specified (legacy format)'
                        })
                    except:
                        continue
            
            return result
            
        except Exception as e:
            m_print(f"Failed to parse LLM response: {e}", stamps=['parse', 'error'])
            return []
    
    def convert_to_global_coords(self, local_anomalies: List[Dict], window_info: Dict) -> List[Dict]:
        """將窗格內座標轉換為全域座標，保留規則信息"""
        global_anomalies = []
        for anomaly in local_anomalies:
            local_row, local_col = anomaly['location']
            global_row = window_info['start_row'] + local_row
            global_col = window_info['start_col'] + local_col
            
            global_anomaly = {
                'location': (global_row, global_col),
                'violated_rule': anomaly['violated_rule'],
                'rule_description': anomaly['rule_description'],
                'window_id': window_info['window_id']
            }
            global_anomalies.append(global_anomaly)
        
        return global_anomalies
    
    def detect_anomalies(self, df: pd.DataFrame, use_openai: bool = False) -> Dict[str, Any]:
        """偵測異常"""
        m_print("Starting anomaly detection...", stamps=['detect'])
        
        # 保存原始DataFrame引用供提示生成和視覺化使用
        self.original_df = df.copy()
        
        # 創建窗格
        windows = self.create_windows(df, self.window_shape)
        
        all_anomalies = []
        
        for i, window_info in enumerate(windows):
            # 跳過純文字區域，只處理數據區域
            if not window_info.get('is_data_area', True):
                m_print(f"Skipping text area window {i+1}/{len(windows)}: {window_info['window_id']}", stamps=['detect'])
                continue
                
            m_print(f"Processing data window {i+1}/{len(windows)}: {window_info['window_id']}", stamps=['detect'])
            
            # 準備提示
            prompt = self.prepare_llm_prompt(window_info['sub_df'], window_info)
            
            # 呼叫LLM
            if use_openai and self.openai_api_key:
                response = self.call_openai(prompt)
            else:
                response = self.call_local_llm(prompt)
            
            m_print(f"LLM response: {response}", stamps=['llm'])
            
            # 解析回應並進行二次驗證
            local_anomalies = self.parse_llm_response(response, df, window_info)
            
            # 轉換為全域座標
            global_anomalies = self.convert_to_global_coords(local_anomalies, window_info)
            
            all_anomalies.extend(global_anomalies)
        
        # 去除重複（基於位置）
        unique_anomalies = []
        seen_locations = set()
        
        for anomaly in all_anomalies:
            location = anomaly['location']
            if location not in seen_locations:
                seen_locations.add(location)
                unique_anomalies.append(anomaly)
        
        # 準備結果
        result = {
            'anomalies': unique_anomalies,
            'abnormal_locs': [anomaly['location'] for anomaly in unique_anomalies],  # 向後兼容
            'total_windows': len(windows),
            'total_anomalies': len(unique_anomalies),
            'rule_violations': self._summarize_rule_violations(unique_anomalies)
        }
        
        m_print(f"Detection completed, found {len(unique_anomalies)} anomaly locations", stamps=['detect'])
        return result
    
    def _summarize_rule_violations(self, anomalies: List[Dict]) -> Dict[int, Dict]:
        """統計違反規則的摘要"""
        violations = {}
        
        for anomaly in anomalies:
            rule_num = anomaly['violated_rule']
            if rule_num > 0:  # 忽略未指定規則的情況
                if rule_num not in violations:
                    violations[rule_num] = {
                        'rule_description': anomaly['rule_description'],
                        'violation_count': 0,
                        'locations': []
                    }
                violations[rule_num]['violation_count'] += 1
                violations[rule_num]['locations'].append(anomaly['location'])
        
        return violations
    
    def visualize_anomalies(self, original_df: pd.DataFrame, anomaly_coords: List[Tuple[int, int]], output_path: str, anomalies_detail: List[Dict] = None):
        """視覺化異常位置，包含規則編號 - 始終顯示完整的原始表格"""
        try:
            # 設定matplotlib中文字體
            if self.font_path:
                font_prop = FontProperties(fname=self.font_path)
                plt.rcParams['font.family'] = font_prop.get_name()
            
            # 創建異常位置到規則的映射
            anomaly_rule_map = {}
            if anomalies_detail:
                for anomaly in anomalies_detail:
                    location = anomaly['location']
                    rule_num = anomaly.get('violated_rule', 0)
                    anomaly_rule_map[location] = rule_num
            
            # 創建圖形 - 基於原始完整DataFrame的大小
            fig, ax = plt.subplots(figsize=(max(12, original_df.shape[1] * 1.5), max(8, original_df.shape[0] * 0.8)))
            
            # 隱藏軸
            ax.set_xlim(0, original_df.shape[1])
            ax.set_ylim(0, original_df.shape[0])
            ax.invert_yaxis()  # 翻轉Y軸使(0,0)在左上角
            
            # 繪製完整的原始表格
            for i in range(original_df.shape[0]):
                for j in range(original_df.shape[1]):
                    # 判斷是否為異常位置
                    is_anomaly = (i, j) in anomaly_coords
                    
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
                    
                    # 添加儲存格內容（來自原始DataFrame）
                    cell_value = original_df.iloc[i, j]
                    if pd.notna(cell_value):
                        text = str(cell_value)
                        if len(text) > 15:  # 稍微增加文字長度限制
                            text = text[:15] + '...'
                        
                        # 如果是異常位置，在文字旁邊添加規則編號
                        if is_anomaly and (i, j) in anomaly_rule_map:
                            rule_num = anomaly_rule_map[(i, j)]
                            if rule_num > 0:
                                # 主要文字（稍微向左）
                                ax.text(j + 0.3, i + 0.5, text, 
                                       ha='center', va='center', 
                                       color=text_color, fontsize=7,
                                       fontproperties=font_prop if self.font_path else None)
                                # 規則編號（右上角，黃色醒目）
                                ax.text(j + 0.8, i + 0.2, f'R{rule_num}', 
                                       ha='center', va='center', 
                                       color='yellow', fontsize=8, fontweight='bold',
                                       bbox=dict(boxstyle="round,pad=0.1", facecolor='black', alpha=0.7),
                                       fontproperties=font_prop if self.font_path else None)
                            else:
                                # 沒有規則編號的異常
                                ax.text(j + 0.5, i + 0.5, text, 
                                       ha='center', va='center', 
                                       color=text_color, fontsize=8,
                                       fontproperties=font_prop if self.font_path else None)
                        else:
                            # 正常儲存格（顯示完整原始內容）
                            ax.text(j + 0.5, i + 0.5, text, 
                                   ha='center', va='center', 
                                   color=text_color, fontsize=8,
                                   fontproperties=font_prop if self.font_path else None)
            
            # 設定標題
            data_areas_processed = len([coord for coord in anomaly_coords if self._is_in_data_area(original_df, coord)])
            title = f'Excel Anomaly Detection Results\n({len(anomaly_coords)} anomalies found, {data_areas_processed} in data areas)'
            plt.title(title, fontproperties=font_prop if self.font_path else None, fontsize=12)
            
            # 添加詳細圖例
            if anomalies_detail:
                legend_lines = [
                    "Legend:",
                    "• Red cells = Anomalies detected",
                    "• R# = Rule number violated (in yellow box)",
                    "• Only numeric data areas were analyzed"
                ]
                legend_text = "\n".join(legend_lines)
                plt.figtext(0.02, 0.02, legend_text, fontsize=9, 
                           fontproperties=font_prop if self.font_path else None,
                           bbox=dict(boxstyle="round,pad=0.3", facecolor='lightgray', alpha=0.8))
            
            # 移除軸標籤
            ax.set_xticks([])
            ax.set_yticks([])
            
            # 儲存圖片
            plt.tight_layout()
            plt.savefig(output_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            m_print(f"Visualization saved to: {output_path} (showing complete original table)", stamps=['visualize'])
            
        except Exception as e:
            m_print(f"Visualization failed: {e}", stamps=['visualize', 'error'])
            LOGger.exception_process(e, stamps=['visualize'])
    
    def _is_in_data_area(self, df: pd.DataFrame, coord: Tuple[int, int]) -> bool:
        """檢查座標是否在數據區域內"""
        row, col = coord
        if row >= df.shape[0] or col >= df.shape[1]:
            return False
        
        # 使用相同的邏輯檢查是否為數據區域
        window_size = 5  # 使用5x1窗格檢查
        start_row = (row // window_size) * window_size
        start_col = col
        end_row = min(start_row + window_size, df.shape[0])
        end_col = min(start_col + 1, df.shape[1])
        
        return self._is_numeric_data_area(df, start_row, start_col, end_row, end_col)

def main():
    parser = argparse.ArgumentParser(description='Excel數據異常偵測器')
    parser.add_argument('file_path', help='要偵測的Excel文件路徑')
    parser.add_argument('--window-shape', nargs=2, type=int, default=[5, 1], 
                       help='窗格大小 (行數 列數)，預設為 5 1')
    parser.add_argument('--use-openai', action='store_true', 
                       help='使用OpenAI而非本地LLM')
    parser.add_argument('--output-dir', default='./anomaly_results', 
                       help='輸出目錄，預設為 ./anomaly_results')
    parser.add_argument('--rules-file', 
                       help='自訂規則文件路徑')
    
    args = parser.parse_args()
    
    # 檢查文件是否存在
    if not os.path.exists(args.file_path):
        print(f"Error: File not found - {args.file_path}")
        return
    
    # 創建輸出目錄
    os.makedirs(args.output_dir, exist_ok=True)
    
    try:
        # 初始化偵測器
        detector = ExcelAnomalyDetector(args.rules_file)
        detector.window_shape = tuple(args.window_shape)
        
        # 載入數據
        df = detector.load_excel_data(args.file_path)
        
        # 偵測異常
        result = detector.detect_anomalies(df, args.use_openai)
        
        # 輸出結果 (避免中文編碼問題)
        try:
            print(f"\n=== Detection Results ===")
            print(f"Total windows: {result['total_windows']}")
            print(f"Anomaly count: {result['total_anomalies']}")
            print(f"Anomaly coordinates: {result['abnormal_locs']}")
            
            # 顯示詳細的規則違反信息
            if result['rule_violations']:
                print(f"\n=== Rule Violations Summary ===")
                for rule_num, violation_info in result['rule_violations'].items():
                    print(f"Rule {rule_num}: {violation_info['rule_description']}")
                    print(f"  Violations: {violation_info['violation_count']}")
                    print(f"  Locations: {violation_info['locations']}")
            
            # 顯示每個異常的詳細信息
            if result['anomalies']:
                print(f"\n=== Detailed Anomaly Information ===")
                for i, anomaly in enumerate(result['anomalies'], 1):
                    location = anomaly['location']
                    rule_num = anomaly['violated_rule']
                    rule_desc = anomaly['rule_description']
                    window_id = anomaly['window_id']
                    
                    print(f"{i}. Location: {location}")
                    print(f"   Violated Rule: {rule_num}")
                    print(f"   Rule Description: {rule_desc}")
                    print(f"   Window: {window_id}")
                    
        except UnicodeEncodeError:
            print(f"\n=== Detection Results ===")
            print(f"Total windows: {result['total_windows']}")
            print(f"Anomaly count: {result['total_anomalies']}")
            print(f"Anomaly coordinates: {result['abnormal_locs']}")
        
        # 儲存結果到JSON
        result_file = os.path.join(args.output_dir, 'anomaly_result.json')
        with open(result_file, 'w', encoding='utf-8') as f:
            # 轉換tuple為list以便JSON序列化
            json_result = result.copy()
            json_result['abnormal_locs'] = [list(loc) for loc in result['abnormal_locs']]
            if 'anomalies' in json_result:
                for anomaly in json_result['anomalies']:
                    anomaly['location'] = list(anomaly['location'])
            if 'rule_violations' in json_result:
                for rule_num, violation_info in json_result['rule_violations'].items():
                    violation_info['locations'] = [list(loc) for loc in violation_info['locations']]
            
            json.dump(json_result, f, ensure_ascii=False, indent=2)
        
        try:
            print(f"\nResults saved to: {result_file}")
        except UnicodeEncodeError:
            print(f"\nResults saved to: {result_file}")
        
        # 視覺化（使用原始完整DataFrame）
        if result['abnormal_locs']:
            output_image = os.path.join(args.output_dir, 'anomaly_visualization.png')
            # 使用保存的原始DataFrame進行視覺化
            original_df = getattr(detector, 'original_df', df)
            detector.visualize_anomalies(original_df, result['abnormal_locs'], output_image, result.get('anomalies', []))
        else:
            try:
                print("No anomalies found, skipping visualization")
            except UnicodeEncodeError:
                print("No anomalies found, skipping visualization")
        
    except Exception as e:
        print(f"Execution failed: {e}")
        LOGger.exception_process(e, stamps=['main'])

if __name__ == '__main__':
    main()
