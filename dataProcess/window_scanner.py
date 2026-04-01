#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
窗格掃描器 - 用指定窗格大小掃描DataFrame
"""

import pandas as pd
import numpy as np
from typing import List, Dict, Tuple, Any, Optional
import argparse
from pathlib import Path
import json
import os

class WindowScanner:
    """DataFrame窗格掃描器"""
    
    def __init__(self, window_shape: Tuple[int, int] = (5, 1)):
        """
        初始化窗格掃描器
        
        Args:
            window_shape: 窗格形狀 (rows, cols)，預設(5, 1)
        """
        self.window_shape = window_shape
        self.window_height, self.window_width = window_shape
        
    def scan_dataframe(self, df: pd.DataFrame, step_size: Tuple[int, int] = (1, 1)) -> List[Dict[str, Any]]:
        """
        掃描DataFrame，產生窗格子dataframe
        
        Args:
            df: 要掃描的DataFrame
            step_size: 步長 (row_step, col_step)，預設(1, 1)
        
        Returns:
            List[Dict]: 包含窗格資訊的列表
        """
        windows = []
        row_step, col_step = step_size
        
        print(f"開始掃描DataFrame，形狀: {df.shape}")
        print(f"窗格大小: {self.window_shape}, 步長: {step_size}")

        # 如果視窗大小超過 DataFrame，給警告並自動縮小
        df_rows, df_cols = df.shape
        new_height = min(self.window_height, df_rows) if df_rows > 0 else self.window_height
        new_width = min(self.window_width, df_cols) if df_cols > 0 else self.window_width
        if new_height != self.window_height or new_width != self.window_width:
            print(
                f"警告: 視窗大小 {self.window_shape} 超過 DataFrame 尺寸 {df.shape}，"
                f"已自動縮小為 ({new_height}, {new_width})"
            )
            self.window_height = new_height
            self.window_width = new_width
            self.window_shape = (new_height, new_width)
        
        # 計算可能的窗格位置
        max_row = df.shape[0] - self.window_height + 1
        max_col = df.shape[1] - self.window_width + 1
        
        window_count = 0
        
        for start_row in range(0, max_row, row_step):
            for start_col in range(0, max_col, col_step):
                end_row = start_row + self.window_height
                end_col = start_col + self.window_width
                
                # 提取窗格
                window_df = df.iloc[start_row:end_row, start_col:end_col].copy()
                
                # 檢查窗格是否有效（不全為空值）
                if self._is_valid_window(window_df):
                    # start_loc = self._get_start_loc_row_index(window_df)
                    start_loc = start_row
                    start_loc_row_name = window_df.index[0]
                    start_loc_row_indicated = [str(idx) for idx in window_df.index.tolist()]
                    
                    window_info = {
                        'window_id': window_count,
                        'position': {
                            'start_row': start_row,
                            'end_row': end_row - 1,  # 包含端點
                            'start_col': start_col,
                            'end_col': end_col - 1   # 包含端點
                        },
                        'shape': window_df.shape,
                        'dataframe': window_df,
                        'values': window_df.values.tolist(),
                        'index_names': window_df.index.tolist(),
                        'column_names': window_df.columns.tolist(),
                        'excel_range': self._get_excel_range(start_row, start_col, end_row-1, end_col-1),
                        'start_loc': start_loc,
                        'start_loc_row_name': start_loc_row_name,
                        'start_loc_row_indicated': start_loc_row_indicated
                    }
                
                windows.append(window_info)
                window_count += 1
        
        print(f"掃描完成，共產生 {len(windows)} 個有效窗格")
        return windows
    
    
    def _is_valid_window(self, window_df: pd.DataFrame) -> bool:
        """
        檢查窗格是否有效
        
        Args:
            window_df: 窗格DataFrame
            
        Returns:
            bool: 是否有效
        """
        # 檢查是否有非空值
        non_null_count = window_df.count().sum()
        total_cells = window_df.shape[0] * window_df.shape[1]
        
        # 至少要有50%的非空值
        return non_null_count >= (total_cells * 0.5)
    
    def _get_excel_range(self, start_row: int, start_col: int, end_row: int, end_col: int) -> str:
        """
        轉換為Excel範圍表示法
        
        Args:
            start_row, start_col, end_row, end_col: 範圍座標 (0-based)
            
        Returns:
            str: Excel範圍字串，如 "A1:A5"
        """
        def col_to_excel(col_idx):
            """將列索引轉換為Excel列名"""
            if col_idx < 26:
                return chr(ord('A') + col_idx)
            else:
                # 處理超過Z的情況 (AA, AB, etc.)
                first_char = chr(ord('A') + (col_idx // 26) - 1)
                second_char = chr(ord('A') + (col_idx % 26))
                return first_char + second_char
        
        start_excel = f"{col_to_excel(start_col)}{start_row + 1}"  # Excel是1-based
        end_excel = f"{col_to_excel(end_col)}{end_row + 1}"
        
        if start_excel == end_excel:
            return start_excel
        else:
            return f"{start_excel}:{end_excel}"
    
    def filter_windows_by_start_loc_row(self, windows: List[Dict[str, Any]], 
                                   start_loc_row_name: Any) -> List[Dict[str, Any]]:
        """
        根據起始列名稱（window 第一列）過濾窗格
        
        Args:
            windows: 窗格列表
            start_loc_row_name: 起始列名稱（window 第一列）
            
        Returns:
            List[Dict]: 符合條件的窗格列表
        """
        filtered_windows = []
        
        for window in windows:
            window_df = window['dataframe']
            
            # 使用 window 第一列作為 start_loc_row
            start_loc_row_idx = 0
            
            if start_loc_row_idx < len(window_df.index):
                current_start_loc_row = window_df.index[start_loc_row_idx]
                
                if current_start_loc_row == start_loc_row_name:
                    filtered_windows.append(window)
        
        print(f"根據start_loc_row_name={start_loc_row_name}過濾，找到 {len(filtered_windows)} 個符合的窗格")
        return filtered_windows
    
    def generate_prompt_data(self, window: Dict[str, Any]) -> Dict[str, Any]:
        """
        為窗格生成提示資料
        
        Args:
            window: 窗格資訊
            
        Returns:
            Dict: 提示資料
        """
        window_df = window['dataframe']
        
        # 提取數值矩陣與名稱（row/column）
        row_names = [str(idx) for idx in window_df.index.tolist()]
        column_names = [str(col) for col in window_df.columns.tolist()]
        values_matrix = []
        numeric_count = 0

        def _coerce_numeric(v):
            if pd.isna(v):
                return None
            if isinstance(v, (int, float, np.number)):
                return float(v)
            try:
                return float(str(v).replace(',', '').strip())
            except Exception:
                return None

        for r in range(len(window_df.index)):
            row_vals = []
            for c in range(len(window_df.columns)):
                v = _coerce_numeric(window_df.iloc[r, c])
                if isinstance(v, (int, float)):
                    numeric_count += 1
                row_vals.append(v)
            values_matrix.append(row_vals)

        prompt_text = (
            f"row_names: {row_names}\n"
            f"column_names: {column_names}\n"
            f"values_matrix: {values_matrix}"
        )
        
        return {
            'values': values_matrix,  # 2D matrix aligned with window (rows x cols)
            'row_names': row_names,
            'index_names': row_names,  # backward-compatible alias
            'column_names': column_names,
            'has_numeric': numeric_count >= 2,
            'prompt_text': prompt_text,
            'window_info': {
                'id': window['window_id'],
                'position': window['position'],
                'excel_range': window['excel_range']
            }
        }

def main():
    """主函數"""
    parser = argparse.ArgumentParser(description='DataFrame窗格掃描器')
    parser.add_argument('excel_file', nargs='?', default='correct_simple.xlsx', 
                       help='Excel檔案路徑（預設: correct_simple.xlsx）')
    parser.add_argument('--window-height', type=int, default=5, 
                       help='窗格高度（預設: 5）')
    parser.add_argument('--window-width', type=int, default=1, 
                       help='窗格寬度（預設: 1）')
    parser.add_argument('--step-row', type=int, default=1, 
                       help='行步長（預設: 1）')
    parser.add_argument('--step-col', type=int, default=1, 
                       help='列步長（預設: 1）')
    parser.add_argument('--start-loc-row-name', type=str, 
                       help='過濾特定起始位置行名稱的窗格')
    
    args = parser.parse_args()
    
    # 檢查檔案路徑
    excel_path = Path(args.excel_file)
    if not excel_path.is_absolute():
        # 相對路徑，從上層目錄尋找
        base_path = Path(__file__).parent.parent.parent
        excel_path = base_path / args.excel_file
    
    if not excel_path.exists():
        print(f"❌ 找不到Excel檔案: {excel_path}")
        return
    
    print(f"使用Excel檔案: {excel_path}")
    
    try:
        # 讀取Excel檔案（需要先用universal_table_detector分析）
        print("需要先使用universal_table_detector.py分析表格結構")
        
        # 簡化處理：直接讀取為DataFrame
        df = pd.read_excel(excel_path, header=0, index_col=0)
        print(f"DataFrame形狀: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print(f"Index: {list(df.index)}")
        
        # 初始化掃描器
        window_shape = (args.window_height, args.window_width)
        scanner = WindowScanner(window_shape)
        
        # 掃描DataFrame
        step_size = (args.step_row, args.step_col)
        windows = scanner.scan_dataframe(df, step_size)
        
        # 過濾特定起始位置行
        if args.start_loc_row_name:
            windows = scanner.filter_windows_by_start_loc_row(windows, args.start_loc_row_name)
        
        # 顯示結果
        print(f"\n=== 掃描結果 ===")
        for i, window in enumerate(windows[:5]):  # 只顯示前5個
            print(f"\n窗格 {window['window_id']}:")
            print(f"  位置: {window['position']}")
            print(f"  Excel範圍: {window['excel_range']}")
            print(f"  形狀: {window['shape']}")
            print(f"  DataFrame:")
            print(window['dataframe'])
            
            # 生成提示資料
            prompt_data = scanner.generate_prompt_data(window)
            print(f"  提示文字: {prompt_data['prompt_text']}")
        
        if len(windows) > 5:
            print(f"\n... 還有 {len(windows) - 5} 個窗格")
        
        # 保存結果
        output_dir = Path(__file__).parent
        output_file = output_dir / f"windows_scan_result_{window_shape[0]}x{window_shape[1]}.json"
        
        # 準備保存的資料（排除DataFrame物件）
        save_data = []
        for window in windows:
            save_window = window.copy()
            save_window['dataframe'] = window['dataframe'].to_dict('records')  # 轉換為可序列化格式
            save_data.append(save_window)
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(save_data, f, ensure_ascii=False, indent=2, default=str)
        
        print(f"\n結果已保存到: {output_file}")
        
    except Exception as e:
        print(f"❌ 處理失敗: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
