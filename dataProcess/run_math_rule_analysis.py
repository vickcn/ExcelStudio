#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
數學規律分析主程式：掃描表格視窗並以 LLM 偵測規律。
"""

import os
import json
import pandas as pd
import numpy as np
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
import argparse
from datetime import datetime as dt
from collections import defaultdict
import random
from concurrent.futures import ThreadPoolExecutor

from utils import LOGger, normalize_idx, repl_names, interval_overlap_seconds

m_logfile = Path(__file__).resolve().parent.parent / "outputs" / "run_math_rule_analysis.log"
m_print = LOGger.addloger(logfile=str(m_logfile))
m_info = lambda *args, **kwargs: None

from window_scanner import WindowScanner
from llm_math_rule_detector import LLMMathRuleDetector
from llm_math_rule_detector import is_dollar_expr
from universal_table_detector import UniversalTableDetector
from rule_audit_normalize import normalize_equation
from rule_degeneracy_filters import (
    should_reject_passed_rule,
    DEFAULT_MIN_NONZERO_COUNT,
    DEFAULT_MIN_DISTINCT_NONZERO,
)


def _with_file_suffix(path: Path, file_suffix: str) -> Path:
    """在副檔名前附加後綴（例如 observed + _phase1 -> observed_phase1.json）。"""
    if not file_suffix:
        return path
    return path.with_name(f"{path.stem}{file_suffix}{path.suffix}")



class MathRuleAnalyzer:
    """Math rule analysis pipeline."""
    
    def __init__(
            self,
            use_openai: bool = False,
            window_shape: tuple = (5, 1),
            openai_model: str = "gpt35_chat",
            consistency_threshold: float = 0.8,
            degeneracy_min_nonzero_count: int = DEFAULT_MIN_NONZERO_COUNT,
            degeneracy_min_distinct_nonzero: int = DEFAULT_MIN_DISTINCT_NONZERO,
            quick_scan_threshold: int = 3,
            quick_scan_seed: Optional[int] = None,
            use_phase1_global_llm_batch: bool = True,
            phase2_overlap_phase1_retro: bool = True,
            step_size: Tuple[int, int] = (1, 1),
        ):
        """
        Args:
            use_openai: 是否使用 OpenAI provider（否則走 TextProcessor 預設模型）
            window_shape: 掃描視窗大小 (rows, cols)
            openai_model: OpenAI 模型別名
            consistency_threshold: 回溯驗證 passed_rules 的一致性門檻（與 CLI 預設一致）
            degeneracy_min_nonzero_count: 退化過濾：每個 $(r,c) 至少需幾筆非零樣本
            degeneracy_min_distinct_nonzero: 退化過濾：非零值至少幾個相異值
            quick_scan_threshold: A線；每 start_loc 群 Phase1 僅抽樣本參數這麼多個視窗做 LLM，同群其餘延至 Phase2（0=關閉 A線）
            quick_scan_seed: 可選；固定後每群 shuffle 可重現
            use_phase1_global_llm_batch: True 時走 Phase1 全域湊滿 batch 排程（預設）
            phase2_overlap_phase1_retro: A 線時是否讓 Phase2（I/O）與 Phase1 回溯+checkpoint（CPU/寫檔）並行
        """
        self.use_openai = use_openai
        self.window_shape = window_shape
        self.openai_model = openai_model
        self.consistency_threshold = consistency_threshold
        self.degeneracy_min_nonzero_count = degeneracy_min_nonzero_count
        self.degeneracy_min_distinct_nonzero = degeneracy_min_distinct_nonzero
        self.quick_scan_threshold = quick_scan_threshold
        self.quick_scan_seed = quick_scan_seed
        self.use_phase1_global_llm_batch = use_phase1_global_llm_batch
        self.phase2_overlap_phase1_retro = phase2_overlap_phase1_retro
        self.step_size = step_size
        
        # Initialize core components
        self.window_scanner = WindowScanner(window_shape)
        self.table_detector = UniversalTableDetector()
        self.llm_detector = LLMMathRuleDetector(prefer_local=not use_openai, openai_model=openai_model)
        
        # 分析結果與計時資訊
        self.analysis_results = []
        self.timer = {'table_detection': 0, 'dataframe_extraction': 0, 'window_scanning': 0, 'llm_analysis': 0, 'retrospective_validation': 0}
        
        m_print("MathRuleAnalyzer initialized")
        m_print(f"step_size: {self.step_size}")
        m_print(f"LLM模式: {'TextProcessor OpenAI' if use_openai else 'TextProcessor remote8b'}")
        m_print(f"視窗大小: {window_shape}")
    
    def _analyze_table_results(
        self,
        table_results: List[Dict[str, Any]],
        *,
        excel_path: Optional[Path] = None,
        full_df: Optional[pd.DataFrame] = None,
        start_loc_row_name: Any = None,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        分析表格結果。
        
        Args:
            table_results: 表格結果
        """
        self.timer['table_detection'] = dt.now()
        if not table_results:
            m_print("未偵測到任何表格")
            return {'success': False, 'error': '未偵測到任何表格'}
        
        m_print(f"tables_found={len(table_results)}")
        
        # 步驟2: 擷取 DataFrame
        m_print("=== 步驟2: 擷取 DataFrame ===")
        if full_df is not None:
            dataframes = self.table_detector.extract_dataframes_from_full(full_df, table_results)
        elif excel_path is not None:
            dataframes = self.table_detector.extract_dataframes(str(excel_path), table_results)
        else:
            return {'success': False, 'error': '缺少 excel_path 或 full_df'}
        self.timer['dataframe_extraction'] = dt.now()
        if not dataframes:
            m_print("無法擷取 DataFrame")
            return {'success': False, 'error': '無法擷取 DataFrame'}
        
        # 以第一個 DataFrame 作為主要分析對象
        main_df = dataframes[0]
        m_print(f"主要 DataFrame 大小: {main_df.shape}")
        m_print(f"Columns: {list(main_df.columns)}")
        m_print(f"Index: {list(main_df.index)}")
        
        # 步驟3: 視窗掃描
        m_print("=== 步驟3: 視窗掃描 ===")
        windows = self.window_scanner.scan_dataframe(main_df, self.step_size)
        self.timer['window_scanning'] = dt.now()
        if not windows:
            m_print("未掃描到任何視窗")
            return {'success': False, 'error': '未掃描到任何視窗'}
        
        # 若指定 start_loc_row_name，則先過濾視窗
        if start_loc_row_name is not None:
            filtered_windows = self.window_scanner.filter_windows_by_start_loc_row(windows, start_loc_row_name)
            m_print(f"filtered_windows={len(filtered_windows)} for start_loc_row_name={start_loc_row_name}")
            windows = filtered_windows
        self.timer['window_scanning'] = dt.now()
        # 步驟4: LLM 規律分析
        m_print("=== 步驟4: LLM 規律分析 ===")
        window_results = []
        valid_rules_count = 0
        total_rules_count = 0
        self.timer['llm_analysis'] = dt.now()
        m_print(f"total_windows={len(windows)}")
        
        prepared_windows = []
        for window in windows:
            prompt_data = self.window_scanner.generate_prompt_data(window)
            if not prompt_data.get('has_numeric'):
                m_print(f"  視窗 {window['window_id']} 無有效數值，略過")
                continue
            prepared_windows.append((window, prompt_data))

        deferred_windows_total = 0
        deferred_windows_skipped_duplicate = 0
        start_loc_row_validation_phase1: Optional[Dict[str, Any]] = None
        pipeline_timing: Optional[Dict[str, Any]] = None

        if self.quick_scan_threshold <= 0:
            batch_inputs = []
            for window, prompt_data in prepared_windows:
                batch_inputs.append({
                    'values_matrix': prompt_data.get('values', []),
                    'row_names': prompt_data.get('row_names', []),
                    'column_names': prompt_data.get('column_names', [])
                })
            batch_results = self.llm_detector.detect_math_rules_batch(
                batch_inputs,
                use_openai=self.use_openai
            )
            self.timer['llm_analysis'] = dt.now()
            for i, (window, prompt_data) in enumerate(prepared_windows):
                llm_result = self._resolve_llm_after_batch(batch_results, i, window, prompt_data)
                wr = self._finalize_window_from_llm(
                    window,
                    prompt_data,
                    llm_result,
                    f"{i + 1}/{len(prepared_windows)}",
                )
                if wr:
                    window_results.append(wr)
                    total_rules_count += wr['validation']['total_rules']
                    valid_rules_count += wr['validation']['valid_rules']
        else:
            if self.quick_scan_seed is not None:
                random.seed(self.quick_scan_seed)
            m_print(
                f"A線 quick_scan_threshold={self.quick_scan_threshold} "
                f"(每 start_loc 群 Phase1 抽樣視窗數；其餘 Phase2；0=關閉)"
            )
            # Phase1：False=逐群各自切 batch；True=全群 Phase1 視窗併成一串再切 batch（見各函式頂註解）。
            if self.use_phase1_global_llm_batch:
                (
                    window_results,
                    valid_rules_count,
                    total_rules_count,
                    deferred_queue,
                ) = self._phase1_quick_scan_global_batch_scheduler(prepared_windows)
            else:
                (
                    window_results,
                    valid_rules_count,
                    total_rules_count,
                    deferred_queue,
                ) = self._phase1_quick_scan_per_group_batches(prepared_windows)

            deferred_windows_total = len(deferred_queue)
            m_print(f"A線 Phase1 完成，deferred_windows={deferred_windows_total}")

            phase1_wr_ids = {wr['window_info']['window_id'] for wr in window_results}
            retro_phase1_list = list(window_results)
            for window, prompt_data in prepared_windows:
                if window['window_id'] not in phase1_wr_ids:
                    retro_phase1_list.append(
                        self._stub_window_result_for_phase1_retro(window, prompt_data)
                    )
            per_loc_known_pre = defaultdict(set)
            for wr in window_results:
                sl = wr['window_info'].get('start_loc')
                per_loc_known_pre[sl] |= self._fingerprints_valid_rules(
                    wr['llm_result'], wr['validation']
                )

            dqc = list(deferred_queue)
            use_parallel_p2 = self.phase2_overlap_phase1_retro and len(dqc) > 0
            executor = None
            phase2_future = None
            if use_parallel_p2:
                self.timer['phase2_start'] = dt.now()
                plk_for_thread = defaultdict(set)
                for k, v in per_loc_known_pre.items():
                    plk_for_thread[k] = set(v)
                executor = ThreadPoolExecutor(max_workers=1)
                phase2_future = executor.submit(
                    self._run_phase2_deferred_batches,
                    dqc,
                    plk_for_thread,
                )

            phase1_retro_start = dt.now()
            self.timer['phase1_retro_start'] = phase1_retro_start
            m_print(
                "=== A線 Phase1 回溯驗證（候選規律 vs 該 start_loc 全 prepared 視窗）==="
            )
            start_loc_row_validation_phase1 = self._retrospective_validate_start_loc_rows(
                retro_phase1_list
            )
            self.timer['phase1_end'] = dt.now()

            # Phase1 checkpoint：先產生一次中間成果（JSON + timer），供長流程先行檢視。
            phase1_success_rate = (valid_rules_count / total_rules_count) if total_rules_count > 0 else 0
            phase1_textprocessor_cost_summary = self._collect_textprocessor_costs(window_results)
            phase1_analysis_result = {
                'success': True,
                'excel_file': str(excel_path),
                'analysis_time': dt.now().isoformat(),
                'analysis_stage': 'phase1',
                'table_detection': {
                    'tables_found': len(table_results),
                    'main_dataframe_shape': main_df.shape,
                    'columns': list(main_df.columns),
                    'index': list(main_df.index)[:10]
                },
                'window_scanning': {
                    'total_windows': len(windows),
                    'window_shape': self.window_shape,
                    'start_loc_row_filter': start_loc_row_name
                },
                'llm_analysis': {
                    'analyzed_windows': len(window_results),
                    'total_rules_found': total_rules_count,
                    'valid_rules_found': valid_rules_count,
                    'overall_success_rate': phase1_success_rate,
                    'llm_mode': 'TextProcessor OpenAI' if self.use_openai else 'TextProcessor remote8b',
                    'quick_scan_threshold': self.quick_scan_threshold,
                    'quick_scan_enabled': self.quick_scan_threshold > 0,
                    'quick_scan_seed': self.quick_scan_seed,
                    'deferred_windows_total': len(deferred_queue),
                    'deferred_windows_skipped_duplicate': deferred_windows_skipped_duplicate,
                    'phase1_global_llm_batch': self.use_phase1_global_llm_batch,
                },
                'start_loc_row_validation': start_loc_row_validation_phase1,
                'start_loc_row_validation_phase1': start_loc_row_validation_phase1,
                'textprocessor_cost_summary': phase1_textprocessor_cost_summary,
                'detailed_results': window_results
            }
            self.save_analysis_snapshot(phase1_analysis_result, file_suffix='_phase1')
            self.save_rule_summary(
                phase1_analysis_result,
                consistency_threshold=self.consistency_threshold,
                file_suffix='_phase1',
                print_summary=False,
            )
            self.save_to_observed_json(phase1_analysis_result, file_suffix='_phase1')
            self.save_timer(file_suffix='_phase1')
            m_print("A線 Phase1 checkpoint 已輸出（_phase1 後綴）")

            phase1_retro_end = dt.now()
            self.timer['phase1_retro_end'] = phase1_retro_end

            phase2_window_results: List[Dict[str, Any]] = []
            va2 = 0
            tb2 = 0
            sk2 = 0
            plk_sync = defaultdict(set)
            for k, v in per_loc_known_pre.items():
                plk_sync[k] = set(v)

            try:
                if use_parallel_p2 and phase2_future is not None:
                    m_print("=== A線 Phase2: 背景補齊 deferred 視窗（與 Phase1 回溯並行中，主線等待結果）===")
                    try:
                        phase2_window_results, va2, tb2, sk2, _ = phase2_future.result()
                    except Exception as ex:
                        m_print(f"Phase2 背景執行失敗，改同步重跑: {ex}")
                        self.timer['phase2_start'] = dt.now()
                        phase2_window_results, va2, tb2, sk2, _ = (
                            self._run_phase2_deferred_batches(dqc, plk_sync)
                        )
                    self.timer['phase2_end'] = dt.now()
                elif len(dqc) > 0:
                    m_print("=== A線 Phase2: 補齊 deferred 視窗（同步）===")
                    self.timer['phase2_start'] = dt.now()
                    phase2_window_results, va2, tb2, sk2, _ = (
                        self._run_phase2_deferred_batches(dqc, plk_sync)
                    )
                    self.timer['phase2_end'] = dt.now()
            finally:
                if executor is not None:
                    executor.shutdown(wait=True)

            p2s = self.timer.get('phase2_start')
            p2e = self.timer.get('phase2_end')
            pipeline_timing = {
                'phase2_parallel_enabled': use_parallel_p2,
                'phase1_retro_start': phase1_retro_start.isoformat(),
                'phase1_retro_end': phase1_retro_end.isoformat(),
                'phase2_start': p2s.isoformat() if isinstance(p2s, dt) else None,
                'phase2_end': p2e.isoformat() if isinstance(p2e, dt) else None,
                'phase2_seconds': (
                    (p2e - p2s).total_seconds()
                    if isinstance(p2s, dt) and isinstance(p2e, dt)
                    else None
                ),
                'overlap_seconds': interval_overlap_seconds(
                    p2s, p2e, phase1_retro_start, phase1_retro_end
                ),
            }

            window_results.extend(phase2_window_results)
            valid_rules_count += va2
            total_rules_count += tb2
            deferred_windows_skipped_duplicate += sk2

        # 計算整體成功率
        overall_success_rate = (valid_rules_count / total_rules_count) if total_rules_count > 0 else 0
        
        # 彙整 TextProcessor 成本資訊
        textprocessor_cost_summary = self._collect_textprocessor_costs(window_results)
        
        # Phase2: use full-group retro validation by stubbing non-LLM windows
        phase2_wr_ids = {wr['window_info']['window_id'] for wr in window_results}
        retro_phase2_list = list(window_results)
        for window, prompt_data in prepared_windows:
            if window['window_id'] not in phase2_wr_ids:
                retro_phase2_list.append(
                    self._stub_window_result_for_phase1_retro(window, prompt_data)
                )
        start_loc_row_validation = self._retrospective_validate_start_loc_rows(retro_phase2_list)
        self.timer['retrospective_validation'] = dt.now()
        source_label = str(excel_path) if excel_path is not None else "<in_memory>"
        analysis_result = {
            'success': True,
            'excel_file': source_label,
            'sheet_name': sheet_name,
            'analysis_time': dt.now().isoformat(),
            'table_detection': {
                'tables_found': len(table_results),
                'main_dataframe_shape': main_df.shape,
                'columns': list(main_df.columns),
                'index': list(main_df.index)[:10]
            },
            'window_scanning': {
                'total_windows': len(windows),
                'window_shape': self.window_shape,
                'start_loc_row_filter': start_loc_row_name
            },
            'llm_analysis': {
                'analyzed_windows': len(window_results),
                'total_rules_found': total_rules_count,
                'valid_rules_found': valid_rules_count,
                'overall_success_rate': overall_success_rate,
                'llm_mode': 'TextProcessor OpenAI' if self.use_openai else 'TextProcessor remote8b',
                'quick_scan_threshold': self.quick_scan_threshold,
                'quick_scan_enabled': self.quick_scan_threshold > 0,
                'quick_scan_seed': self.quick_scan_seed,
                'deferred_windows_total': deferred_windows_total,
                'deferred_windows_skipped_duplicate': deferred_windows_skipped_duplicate,
                'phase1_global_llm_batch': self.use_phase1_global_llm_batch,
                **({'pipeline_timing': pipeline_timing} if pipeline_timing else {}),
            },
            'start_loc_row_validation': start_loc_row_validation,
            'start_loc_row_validation_phase1': start_loc_row_validation_phase1,
            'textprocessor_cost_summary': textprocessor_cost_summary,
            'detailed_results': window_results
        }
        
        m_print(f"\n=== 分析結果 ===")
        m_print(f"analyzed_windows={len(window_results)}")
        m_print(f"rules_total={total_rules_count} rules_valid={valid_rules_count}")
        m_print(f"成功率: {overall_success_rate:.2%}")
        
        # 顯示 TextProcessor 成本資訊
        if textprocessor_cost_summary['has_costs']:
            m_print(f"\n=== TextProcessor 成本資訊 ===")
            m_print(f"呼叫次數: {textprocessor_cost_summary['total_calls']}")
            m_print(f"Token 用量: {textprocessor_cost_summary['total_usage']}")
            if textprocessor_cost_summary['total_billing']:
                m_print(f"計價資訊: {textprocessor_cost_summary['total_billing']}")
        else:
            m_print(f"\n=== TextProcessor 成本資訊 ===")
            m_print("未取得 TextProcessor 成本資料")

        self.timer['analysis_end'] = dt.now()
        return analysis_result

    def analyze_excel_file(self, excel_file: str, start_loc_row_name: Any = None) -> Dict[str, Any]:
        """
        分析 Excel 檔案中的數學規律。
        
        Args:
            excel_file: Excel 檔案路徑
            start_loc_row_name: 指定起始列名稱（可選）
            
        Returns:
            Dict: 分析結果
        """
        self.timer['analysis_start'] = dt.now()
        self.timer['table_detection'] = dt.now()
        excel_path = Path(excel_file)
        if not excel_path.is_absolute():
            # 相對路徑改為以專案根目錄為基準
            base_path = Path(__file__).parent.parent.parent
            excel_path = base_path / excel_file
        self.timer['dataframe_extraction'] = dt.now()
        if not excel_path.exists():
            m_print(f"找不到 Excel 檔案: {excel_path}")
            return {'success': False, 'error': f'找不到 Excel 檔案: {excel_path}'}
        self.timer['window_scanning'] = dt.now()
        m_print(f"開始分析 Excel 檔案: {excel_path}")
        
        try:
            # 步驟1: 偵測表格區塊
            m_print("=== 步驟1: 偵測表格區塊 ===")
            table_results = self.table_detector.detect_tables_by_analysis(
                str(excel_path), 
                table_mode='pure_numeric',
                use_llm=False
            )
            analysis_result = self._analyze_table_results(
                table_results,
                excel_path=excel_path,
                start_loc_row_name=start_loc_row_name,
                sheet_name=None,
            )
            return analysis_result
        except Exception as e:
            m_print(f"分析流程發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            return {'success': False, 'error': str(e)}

    def analyze_sheetTable(
        self,
        data: List[List[str]],
        start_loc_row_name: Any = None,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        分析單一表格（JSON 矩陣）中的數學規律。
        
        Args:
            data: 二維矩陣（表格內容）
            start_loc_row_name: 指定起始列名稱（可選）
            sheet_name: sheet 名稱（可選）
            
        Returns:
            Dict: 分析結果
        """
        self.timer['analysis_start'] = dt.now()
        self.timer['table_detection'] = dt.now()
        self.timer['dataframe_extraction'] = dt.now()
        self.timer['window_scanning'] = dt.now()
        fulltable = pd.DataFrame(data)
        m_print(f"開始分析 母表格: {fulltable.shape}")
        
        try:
            # 步驟1: 偵測表格區塊
            m_print("=== 步驟1: 偵測表格區塊 ===")
            table_results = self.table_detector._detect_tables_from_full(
                fulltable,
                table_mode='pure_numeric',
                use_llm=False,
                prefer_local=False,
            )
            analysis_result = self._analyze_table_results(
                table_results,
                full_df=fulltable,
                start_loc_row_name=start_loc_row_name,
                sheet_name=sheet_name,
            )
            return analysis_result
        except Exception as e:
            m_print(f"分析流程發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            return {'success': False, 'error': str(e)}
    
    def _postprocess_llm_equations(self, llm_result: Dict[str, Any], prompt_data: Dict[str, Any]) -> None:
        """對 llm_result['rules'] 做 equation_sides / equation 與列名對齊（原地修改）。"""
        llm_result['_row_names'] = prompt_data.get('row_names', [])
        llm_result['_column_names'] = prompt_data.get('column_names', [])
        llm_result['_index_names'] = prompt_data.get('index_names', [])

        name_to_index = {str(nm): i for i, nm in enumerate(prompt_data.get('row_names', []))}

        for r in llm_result.get('rules', []):
            sides = r.get('equation_sides')
            eq = r.get('equation')
            if not (isinstance(sides, list) and len(sides) == 2 and all(is_dollar_expr(x) for x in sides)):
                if isinstance(eq, str) and '=' in eq and '$' in eq:
                    L, R = [t.strip() for t in eq.split('=', 1)]
                    r['equation_sides'] = [normalize_idx(L), normalize_idx(R)]
                elif isinstance(sides, list) and len(sides) == 2:
                    L, R = repl_names(sides[0], name_to_index), repl_names(sides[1], name_to_index)
                    if is_dollar_expr(L) and is_dollar_expr(R):
                        r['equation_sides'] = [normalize_idx(L), normalize_idx(R)]
            if isinstance(r.get('equation_sides'), list) and len(r['equation_sides']) == 2:
                r['equation'] = f"{r['equation_sides'][0]} = {r['equation_sides'][1]}"

    def _resolve_llm_after_batch(
            self,
            batch_results: List[Any],
            batch_index: int,
            window: Dict[str, Any],
            prompt_data: Dict[str, Any],
        ) -> Optional[Dict[str, Any]]:
        llm_result = batch_results[batch_index] if batch_index < len(batch_results) else None
        if not llm_result:
            m_print("  /chat/batch 無結果，回退到單筆 /chat")
            llm_result = self.llm_detector.detect_math_rules(
                prompt_data.get('values', []),
                prompt_data.get('row_names', []),
                prompt_data.get('column_names', []),
                self.use_openai
            )
        return llm_result

    def _finalize_window_from_llm(
            self,
            window: Dict[str, Any],
            prompt_data: Dict[str, Any],
            llm_result: Optional[Dict[str, Any]],
            progress_hint: str,
        ) -> Optional[Dict[str, Any]]:
        """若 LLM 有結果則後處理、validate、組 window_result；否則回傳 None。"""
        m_print(f"\n分析視窗 {window['window_id']} ({progress_hint})")
        self.timer['llm_analysis'] = dt.now()
        if not llm_result:
            m_print("  LLM 分析失敗")
            return None
        self._postprocess_llm_equations(llm_result, prompt_data)
        validation = self.llm_detector.validate_rules(
            prompt_data['values'],
            llm_result.get('rules', [])
        )
        window_result = {
            'window_info': {
                **window,
                'start_loc_row_name': window.get('start_loc_row_name', 'unknown')
            },
            'prompt_data': prompt_data,
            'llm_result': llm_result,
            'validation': validation
        }
        m_print(f"  rules={validation['total_rules']} valid={validation['valid_rules']}")
        return window_result

    def _fingerprints_valid_rules(self, llm_result: Dict[str, Any], validation: Dict[str, Any]) -> set:
        """通過 validate 的規律以 normalize_equation 去重後的集合。"""
        rules = llm_result.get('rules', [])
        details = validation.get('validation_details', [])
        out = set()
        for i, det in enumerate(details):
            if not det.get('is_valid'):
                continue
            if i >= len(rules):
                continue
            eq = rules[i].get('equation', '')
            fp = normalize_equation(str(eq) if eq else '')
            if fp:
                out.add(fp)
        return out

    def _stub_window_result_for_phase1_retro(
        self, window: Dict[str, Any], prompt_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Phase1 尚未做 LLM 的視窗：僅帶 prompt_data，供回溯在每個視窗上重驗證 equation。
        """
        return {
            'window_info': {
                **window,
                'start_loc_row_name': window.get('start_loc_row_name', 'unknown'),
            },
            'prompt_data': prompt_data,
            'llm_result': {'rules': []},
            'validation': {
                'total_rules': 0,
                'valid_rules': 0,
                'success_rate': 0.0,
                'validation_details': [],
            },
            'retro_stub_phase1': True,
        }

    def _run_phase2_deferred_batches(
        self,
        deferred_queue: List[Any],
        per_loc_known: Dict[Any, Any],
    ) -> Tuple[List[Dict[str, Any]], int, int, int, Dict[Any, set]]:
        """
        A 線 Phase2：將 deferred 視窗以 batch LLM 補齊，並依 start_loc 已知規律指紋做聯集去重略過。

        回傳 (phase2 新增的 window_results, valid_rules 增量, total_rules 增量,
        略過重複視窗計數, 更新後的 per_loc 指紋映射)。
        """
        dq = list(deferred_queue)
        plk: Dict[Any, set] = defaultdict(set)
        for k, v in per_loc_known.items():
            plk[k] = set(v)
        phase2_results: List[Dict[str, Any]] = []
        va = 0
        tb = 0
        sk = 0
        while dq:
            n = min(100, len(dq))
            chunk = dq[:n]
            dq = dq[n:]
            batch_inputs = [{
                'values_matrix': pd_item.get('values', []),
                'row_names': pd_item.get('row_names', []),
                'column_names': pd_item.get('column_names', [])
            } for _, pd_item in chunk]
            batch_results = self.llm_detector.detect_math_rules_batch(
                batch_inputs,
                use_openai=self.use_openai
            )
            self.timer['llm_analysis'] = dt.now()
            for i, (window, prompt_data) in enumerate(chunk):
                llm_result = self._resolve_llm_after_batch(batch_results, i, window, prompt_data)
                wr = self._finalize_window_from_llm(
                    window,
                    prompt_data,
                    llm_result,
                    f"Phase2 {i + 1}/{len(chunk)}",
                )
                if not wr:
                    continue
                sl = wr['window_info'].get('start_loc')
                new_fps = self._fingerprints_valid_rules(wr['llm_result'], wr['validation'])
                if new_fps and new_fps <= plk[sl]:
                    sk += 1
                    m_print(
                        f"  A線略過重複規律視窗 {window['window_id']} (不寫入 detailed_results)"
                    )
                    continue
                phase2_results.append(wr)
                v = wr['validation']
                tb += v['total_rules']
                va += v['valid_rules']
                plk[sl] |= new_fps
        return phase2_results, va, tb, sk, dict(plk)

    def _phase1_quick_scan_per_group_batches(
        self,
        prepared_windows: List[Any],
    ) -> Tuple[List[Dict[str, Any]], int, int, List[Any]]:
        """
        Phase1（預設路徑）：逐 start_loc 群處理。

        語意變更註記（舊實作已刪）：
        - 已刪除：群內邊跑 LLM 邊累計「相異 validate 通過規律」達 quick_scan_threshold 才 defer 剩餘。
        - 現行：quick_scan_threshold = 每群 shuffle 後**最多幾個視窗**進 Phase1 LLM；
          同群其餘視窗**不經 Phase1**，直接 append 到 deferred_queue，留給 Phase2 批量送問。
        """
        window_results: List[Dict[str, Any]] = []
        valid_rules_count = 0
        total_rules_count = 0
        by_loc: Dict[Any, List] = defaultdict(list)
        for window, prompt_data in prepared_windows:
            by_loc[window.get('start_loc')].append((window, prompt_data))
        deferred_queue: List[Any] = []
        phase1_batch_max = 100

        for loc in sorted(by_loc.keys(), key=lambda k: str(k)):
            pending = list(by_loc[loc])
            random.shuffle(pending)
            k = min(self.quick_scan_threshold, len(pending))
            phase1_part = pending[:k]
            deferred_queue.extend(pending[k:])

            for st in range(0, len(phase1_part), phase1_batch_max):
                chunk = phase1_part[st:st + phase1_batch_max]
                batch_inputs = [{
                    'values_matrix': pd_item.get('values', []),
                    'row_names': pd_item.get('row_names', []),
                    'column_names': pd_item.get('column_names', [])
                } for _, pd_item in chunk]
                batch_results = self.llm_detector.detect_math_rules_batch(
                    batch_inputs,
                    use_openai=self.use_openai
                )
                self.timer['llm_analysis'] = dt.now()
                for i, (window, prompt_data) in enumerate(chunk):
                    llm_result = self._resolve_llm_after_batch(batch_results, i, window, prompt_data)
                    wr = self._finalize_window_from_llm(
                        window,
                        prompt_data,
                        llm_result,
                        f"Phase1 {i + 1}/{len(chunk)}",
                    )
                    if wr:
                        window_results.append(wr)
                        v = wr['validation']
                        total_rules_count += v['total_rules']
                        valid_rules_count += v['valid_rules']

        return window_results, valid_rules_count, total_rules_count, deferred_queue

    def _phase1_quick_scan_global_batch_scheduler(
        self,
        prepared_windows: List[Any],
    ) -> Tuple[List[Dict[str, Any]], int, int, List[Any]]:
        """
        Phase1（全域 batch）：與 per_group **同一抽樣語意**，差在 LLM 請求前先把「各群 Phase1 視窗」合併再切塊。

        已刪除之舊版（勿復用）：round-robin 依「相異規律數達標」決定 defer，與 quick_scan_threshold 新定義衝突。

        TODO（可交給呼叫端自訂的細節，目前採最簡併序）：
        - phase1_flat 的串接順序：現為 sorted(start_loc) 再依群 append；若需 round-robin 交錯各群 Phase1 視窗以平衡延遲，可在此調整。
        - phase1_batch_max 是否改為與 detect_math_rules_batch 參數或 CLI 一致。
        """
        window_results: List[Dict[str, Any]] = []
        valid_rules_count = 0
        total_rules_count = 0
        deferred_queue: List[Any] = []
        phase1_batch_max = 100

        by_loc: Dict[Any, List] = defaultdict(list)
        for window, prompt_data in prepared_windows:
            by_loc[window.get('start_loc')].append((window, prompt_data))

        phase1_flat: List[Tuple[Dict[str, Any], Dict[str, Any]]] = []
        for loc in sorted(by_loc.keys(), key=lambda k: str(k)):
            pending = list(by_loc[loc])
            random.shuffle(pending)
            k = min(self.quick_scan_threshold, len(pending))
            phase1_flat.extend(pending[:k])
            deferred_queue.extend(pending[k:])

        for st in range(0, len(phase1_flat), phase1_batch_max):
            chunk = phase1_flat[st:st + phase1_batch_max]
            batch_inputs = [{
                'values_matrix': pd_item.get('values', []),
                'row_names': pd_item.get('row_names', []),
                'column_names': pd_item.get('column_names', [])
            } for _, pd_item in chunk]
            batch_results = self.llm_detector.detect_math_rules_batch(
                batch_inputs,
                use_openai=self.use_openai
            )
            self.timer['llm_analysis'] = dt.now()
            for i, (window, prompt_data) in enumerate(chunk):
                llm_result = self._resolve_llm_after_batch(batch_results, i, window, prompt_data)
                wr = self._finalize_window_from_llm(
                    window,
                    prompt_data,
                    llm_result,
                    f"Phase1-Global {i + 1}/{len(chunk)}",
                )
                if wr:
                    window_results.append(wr)
                    v = wr['validation']
                    total_rules_count += v['total_rules']
                    valid_rules_count += v['valid_rules']

        return window_results, valid_rules_count, total_rules_count, deferred_queue

    def _validate_by_start_loc_row(self, window_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        依 start_loc_row 分組，檢查群組內規律一致性。
        
        Args:
            window_results: 視窗分析結果列表
            
        Returns:
            Dict: 每個 start_loc_row 的驗證結果摘要
        """
        # 依 start_loc_row_name 分組
        start_loc_row_groups = defaultdict(list)
        
        for result in window_results:
            if 'window_info' in result:
                start_loc_row = result['window_info'].get('start_loc_row_name', 'unknown')
                start_loc_row_groups[start_loc_row].append(result)
        
        m_print(f"\n=== Start Row 分組驗證 ===")
        m_print(f"共 {len(start_loc_row_groups)} 個 start_loc_row")
        
        validation_summary = {}
        total_groups = len(start_loc_row_groups)
        valid_groups = 0
        
        for start_loc_row, group_results in start_loc_row_groups.items():
            m_print(f"\nStart Row: {start_loc_row}")
            m_print(f"  視窗數量: {len(group_results)}")
            
            # 蒐集群組內所有規律
            all_rules = []
            for result in group_results:
                if 'llm_result' in result and 'rules' in result['llm_result']:
                    all_rules.extend(result['llm_result']['rules'])
            
            # 檢查規律一致性
            consistent_rules = self._check_rule_consistency(group_results)
            
            group_validation = {
                'window_count': len(group_results),
                'total_rules': len(all_rules),
                'consistent_rules': len(consistent_rules),
                'consistency_rate': len(consistent_rules) / len(all_rules) if all_rules else 0,
                'is_valid_group': len(consistent_rules) > 0
            }
            
            if group_validation['is_valid_group']:
                valid_groups += 1
            
            validation_summary[start_loc_row] = group_validation
            
            m_print(f"  規律總數: {group_validation['total_rules']}")
            m_print(f"  一致規律數: {group_validation['consistent_rules']}")
            m_print(f"  一致性比率: {group_validation['consistency_rate']:.2%}")
        
        overall_validation = {
            'total_start_loc_rows': total_groups,
            'valid_start_loc_rows': valid_groups,
            'start_loc_row_success_rate': valid_groups / total_groups if total_groups > 0 else 0,
            'details': validation_summary
        }
        
        m_print(f"\n整體 Start Loc Row 驗證:")
        m_print(f"  start_loc_row 總數: {total_groups}")
        m_print(f"  有效 start_loc_row 數: {valid_groups}")
        m_print(f"  Start Loc Row 成功率: {overall_validation['start_loc_row_success_rate']:.2%}")
        
        return overall_validation
    
    def _retrospective_validate_start_loc_rows(self, window_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        以回溯方式驗證每個 start_loc_row 的規律一致性。
        驗證流程：
        1. 依 start_loc_row_index 分組（key 為 row index）
        2. 蒐集每組中所有候選規律
        3. 逐條規律在該組視窗中回溯驗證
        4. 計算一致性比率與通過門檻規律
        
        Args:
            window_results: 視窗分析結果
        Returns:
            Dict: 回溯驗證結果（key 為 row index）
        """
        # 依 start_loc_row_index 分群，key 為 row index
        start_loc_row_groups = defaultdict(list)
        
        for result in window_results:
            if 'window_info' in result:
                start_loc = result['window_info'].get('start_loc', 'unknown')
                start_loc_row_groups[start_loc].append(result)
        
        m_print(f"\n=== 回溯驗證：依 Start Row 分組 ===")
        m_print(f"共 {len(start_loc_row_groups)} 個 start_loc_row")
        
        validation_summary = {}
        total_groups = len(start_loc_row_groups)
        valid_groups = 0
        
        for start_loc, group_results in start_loc_row_groups.items():
            # 先取此群第一個視窗，帶出 row 名稱與標示資訊
            first_window = group_results[0]
            start_loc_row_name = first_window['window_info'].get('start_loc_row_name', 'unknown')
            start_loc_row_indicated = first_window['window_info'].get('start_loc_row_indicated', [])
            
            m_print(f"\nStart Row Index: {start_loc}")
            m_print(f"  Start Row Name: {start_loc_row_name}")
            m_print(f"  視窗數量: {len(group_results)}")
            
            # 蒐集此 start_loc_row 中所有規律
            all_rules_in_start_loc_row = self._collect_all_rules_in_start_loc_row(group_results)
            m_print(f"  候選規律總數: {len(all_rules_in_start_loc_row)}")
            
            # 逐條規律進行回溯驗證
            rule_consistency_results = self._retrospective_validate_rules_in_start_loc_row(
                all_rules_in_start_loc_row, group_results
            )
            
            # # 舊版：群組整體一致性統計
            # total_windows = len(group_results)
            # overall_consistency_rate = 0.0
            # if rule_consistency_results:
            #     overall_consistency_rate = sum(
            #         result['consistency_rate'] for result in rule_consistency_results.values()
            #     ) / len(rule_consistency_results)
            
            # # 舊版整體驗證統計（保留參考）
            # group_validation = {
            #     'window_count': total_windows,
            #     'total_rules_found': len(all_rules_in_start_loc_row),
            #     'rule_details': rule_consistency_results,
            #     'overall_consistency_rate': overall_consistency_rate,
            #     'is_valid_group': overall_consistency_rate > 0
            # }
            
            # 以每條規律為單位計算是否通過一致性門檻
            total_windows = len(group_results)
            threshold = self.consistency_threshold

            passed_rules = {
                rule_text: details
                for rule_text, details in rule_consistency_results.items()
                if details.get('consistency_rate', 0.0) >= threshold
            }

            passed_rules, degeneracy_filtered_rules = self._apply_degeneracy_filters_to_passed_rules(
                passed_rules, group_results
            )

            group_validation = {
                'start_loc_row_name': start_loc_row_name,
                'start_loc_row_indicated': start_loc_row_indicated,
                'window_count': total_windows,
                'total_rules_found': len(all_rules_in_start_loc_row),
                # 保留完整規律細節，供後續輸出與除錯
                'rule_details': rule_consistency_results,
                # 通過一致性門檻的規律
                'passed_rules': passed_rules,
                # 通過門檻但因退化過濾剔除的規律（策略 1+2，見 rule_degeneracy_filters）
                'degeneracy_filtered_rules': degeneracy_filtered_rules,
                'passed_rules_count': len(passed_rules),
                'passed_rules_rate': (
                    len(passed_rules) / max(1, len(rule_consistency_results))
                    if rule_consistency_results else 0.0
                ),
                # 只要有至少一條規律通過門檻，則此群組有效
                'is_valid_group': len(passed_rules) > 0
            }
            
            if group_validation['is_valid_group']:
                valid_groups += 1
            
            validation_summary[start_loc] = group_validation

            m_print("Rule consistency summary generated")
            sorted_rules = sorted(
                rule_consistency_results.items(),
                key=lambda x: x[1].get('consistency_rate', 0.0),
                reverse=True
            )
            for i, (rule_text, details) in enumerate(sorted_rules):
                m_print(
                    f"    規律{i+1}: {rule_text[:60]}{'...' if len(rule_text) > 60 else ''}"
                )
                m_print(
                    f"      一致性比率: {details.get('consistency_rate', 0.0):.1%} "
                    f"({details.get('valid_windows', 0)}/{details.get('total_windows', 0)})"
                )
        
        overall_validation = {
            'total_start_loc_rows': total_groups,
            'valid_start_loc_rows': valid_groups,
            'start_loc_row_success_rate': valid_groups / total_groups if total_groups > 0 else 0,
            'details': validation_summary,
            'validation_method': 'two_stage_retrospective'
        }
        
        m_print(f"\n回溯驗證總結:")
        m_print(f"  start_loc_row 總數: {total_groups}")
        m_print(f"  有效 start_loc_row 數: {valid_groups}")
        m_print(f"  Start Loc Row 成功率: {overall_validation['start_loc_row_success_rate']:.2%}")
        
        return overall_validation

    def _apply_degeneracy_filters_to_passed_rules(
        self,
        passed_rules: Dict[str, Dict[str, Any]],
        group_results: List[Dict[str, Any]],
    ) -> tuple[Dict[str, Dict[str, Any]], Dict[str, Dict[str, Any]]]:
        """
        對已達 consistency 門檻的 passed_rules 再做退化過濾（策略 1+2）。

        實作位於 rule_degeneracy_filters；此處僅負責逐條呼叫與分流。

        Returns:
            (保留的 passed_rules, 被剔除項及其 degeneracy_reject_reason)
        """
        kept: Dict[str, Dict[str, Any]] = {}
        filtered_out: Dict[str, Dict[str, Any]] = {}
        for rule_text, details in passed_rules.items():
            equation = details.get('rule_equation', '') or ''
            equation_sides = details.get('equation_sides')
            reject, reason = should_reject_passed_rule(
                equation,
                equation_sides,
                group_results,
                min_nonzero_count=self.degeneracy_min_nonzero_count,
                min_distinct_nonzero=self.degeneracy_min_distinct_nonzero,
            )
            if reject:
                out = dict(details)
                out['degeneracy_reject_reason'] = reason
                filtered_out[rule_text] = out
            else:
                kept[rule_text] = details
        return kept, filtered_out
    
    def _collect_textprocessor_costs(self, window_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        彙整 TextProcessor 使用量與計價資訊。
        
        Args:
            window_results: 視窗分析結果列表
            
        Returns:
            Dict: 成本彙整資訊
        """
        total_usage = {
            'prompt_tokens': 0,
            'completion_tokens': 0,
            'total_tokens': 0
        }
        
        total_billing = {}
        total_calls = 0
        cost_details = []
        
        for window_result in window_results:
            if 'llm_result' in window_result:
                llm_result = window_result['llm_result']
                
                # 收集 usage 與 token_summary
                usage = llm_result.get('textprocessor_usage', {})
                token_summary = llm_result.get('textprocessor_token_summary', {})
                
                # 優先使用 token_summary，其次使用 usage
                if token_summary and token_summary.get('total'):
                    total_info = token_summary['total']
                    total_usage['prompt_tokens'] += total_info.get('prompt_tokens', 0)
                    total_usage['completion_tokens'] += total_info.get('completion_tokens', 0)
                    total_usage['total_tokens'] += total_info.get('total_tokens', 0)
                    total_calls += total_info.get('calls', 1)
                    m_print(f" 收集到 token_summary: {total_info}")
                elif usage:
                    total_usage['prompt_tokens'] += usage.get('prompt_tokens', 0)
                    total_usage['completion_tokens'] += usage.get('completion_tokens', 0)
                    total_usage['total_tokens'] += usage.get('total_tokens', 0)
                    total_calls += 1
                    m_print(f" 收集到 usage: {usage}")
                else:
                    m_print(f" 未找到 usage/token_summary，llm_result keys: {list(llm_result.keys())}")
                    # 即使沒有 usage，也記一次呼叫
                    total_calls += 1
                
                # 收集 billing 資訊
                billing = llm_result.get('textprocessor_billing', {})
                if billing:
                    # billing 中數值欄位累加，其他欄位保留最新值
                    for key, value in billing.items():
                        if isinstance(value, (int, float)):
                            total_billing[key] = total_billing.get(key, 0) + value
                        else:
                            total_billing[key] = value
                
                # 收集每個視窗的成本明細
                if usage or billing:
                    cost_details.append({
                        'window_id': window_result.get('window_info', {}).get('window_id', 'unknown'),
                        'usage': usage,
                        'billing': billing,
                        'post_id': llm_result.get('textprocessor_post_id'),
                        'timestamp': llm_result.get('textprocessor_timestamp')
                    })
        
        return {
            'total_usage': total_usage,
            'total_billing': total_billing,
            'total_calls': total_calls,
            'cost_details': cost_details,
            'has_costs': total_calls > 0
        }
    
    def _collect_all_rules_in_start_loc_row(self, group_results: List[Dict[str, Any]]) -> List[str]:
        """
        蒐集指定 start_loc_row 群組中的所有規律方程式。
        Args:
            group_results: 同一個 start_loc_row 的視窗結果
        Returns:
            List[str]: 去重後的規律 equation 清單
        """
        all_rules = set()
        
        for window_result in group_results:
            if 'llm_result' in window_result and 'rules' in window_result['llm_result']:
                for rule in window_result['llm_result']['rules']:
                    # 使用 equation 作為規律的標準鍵值
                    rule_equation = rule.get('equation', '')
                    if rule_equation:
                        normalized_rule = ' '.join(rule_equation.split()).strip()
                        all_rules.add(normalized_rule)
        
        return list(all_rules)
    
    def _retrospective_validate_rules_in_start_loc_row(self, all_rules: List[str], group_results: List[Dict[str, Any]]) -> Dict[str, Dict]:
        """
        對 start_loc_row 內所有規律做回溯驗證。
        
        Args:
            all_rules: 該 start_loc_row 的規律 equation 列表
            group_results: 該 start_loc_row 的視窗結果
        Returns:
            Dict: 各規律的驗證結果細節
        """
        total_windows = len(group_results)
        rule_results = {}
        
        for rule_equation in all_rules:
            valid_windows = 0
            validation_details = []
            
            # 找出此 equation 對應的描述（description）
            rule_description = self._find_description_for_equation(rule_equation, group_results)
            m_info(f"  回溯驗證規律: {rule_equation}")
            m_info(f"  規律描述: {rule_description}")
            
            # 嘗試從任一視窗找到 equation_sides
            equation_sides = None
            found = False

            for window_result in group_results:
                if 'llm_result' in window_result and 'rules' in window_result['llm_result']:
                    for rule in window_result['llm_result']['rules']:
                        rule_eq = rule.get('equation', '')
                        # if len(rule_eq) == len(rule_equation):
                        # compare normalized equation text
                        # if ' '.join(rule_eq.split()).strip() == ' '.join(rule_equation.split()).strip():
                        if normalize_equation(rule_eq) == normalize_equation(rule_equation):
                            eq_sides = rule.get('equation_sides')
                            if eq_sides and isinstance(eq_sides, list) and len(eq_sides) == 2:
                                equation_sides = eq_sides
                                found = True
                                break  # 只有找到有效 sides 才跳出 inner loop
                    if found:
                        break      # inner 有找到才跳出 outer loop
            
            if equation_sides:
                m_info(f"  找到 equation_sides: {equation_sides}")
            else:
                m_print(f"  未找到 equation_sides，將以 equation 直接驗證")
            
            # 在每個視窗上驗證此規律
            for window_result in group_results:
                window_id = window_result.get('window_info', {}).get('window_id', 'unknown')
                
                # 回傳：是否有效、驗證狀態
                is_valid, validation_status = self._check_rule_validation_in_window(rule_equation, window_result)
                
                m_info(f"    視窗{window_id}: {validation_status} -> {'有效' if is_valid else '無效'}")
                
                if is_valid:
                    valid_windows += 1
                
                validation_details.append({
                    'window_id': window_id,
                    'is_valid': is_valid,
                    'rule': rule_equation,
                    'validation_status': validation_status
                })
            
            # 計算一致性比率
            consistency_rate = valid_windows / total_windows
            
            # 以 description 當 key，同時保留 equation
            rule_results[rule_description] = {
                'total_windows': total_windows,
                'valid_windows': valid_windows,
                'invalid_windows': total_windows - valid_windows,
                'consistency_rate': consistency_rate,
                'validation_details': validation_details,
                'rule_description': rule_description,
                'rule_equation': rule_equation,
                'equation_sides': equation_sides  # 可能為 None（未提供 sides）
            }
        
        return rule_results
    
    def _find_description_for_equation(self, equation: str, group_results: List[Dict[str, Any]]) -> str:
        """
        找出 equation 對應的最常見 description。
        
        Args:
            equation: 規律 equation，例如 "$0 - $1 = $2"
            group_results: 視窗結果列表
            
        Returns:
            str: 最常見 description，找不到則回傳 equation
        """
        normalized_equation = ' '.join(equation.split()).strip()
        description_counts = {}
        
        for window_result in group_results:
            if 'llm_result' in window_result and 'rules' in window_result['llm_result']:
                for rule in window_result['llm_result']['rules']:
                    rule_equation = rule.get('equation', '')
                    if rule_equation:
                        rule_normalized = ' '.join(rule_equation.split()).strip()
                        if rule_normalized == normalized_equation:
                            description = rule.get('description', '')
                            if description:
                                description_counts[description] = description_counts.get(description, 0) + 1
        
        # 回傳最常見的 description
        if description_counts:
            most_common_description = max(description_counts.items(), key=lambda x: x[1])[0]
            return most_common_description
        
        # 沒有 description 時，直接回傳 equation
        return equation
    
    def _check_rule_validation_in_window(self, rule_text: str, window_result: Dict[str, Any]) -> tuple[bool, str]:
        """
        檢查規律在單一視窗中的驗證狀態。
        若無既有驗證資料，則改走重驗證流程。
        
        Args:
            rule_text: 規律文字
            window_result: 視窗結果
            
        Returns:
            tuple[bool, str]: (是否有效, 驗證狀態)
            驗證狀態: 'existing_valid', 'existing_invalid', 'revalidated_valid', 'revalidated_invalid', 'skip_invalid'
        """
        if 'validation' not in window_result:
            m_print(f"      視窗{window_result.get('window_info', {}).get('window_id', 'unknown')} 無既有驗證，改做重驗證")
            is_valid = self._validate_rule_with_window_data(rule_text, window_result)
            status = 'revalidated_valid' if is_valid else 'revalidated_invalid'
            return is_valid, status
        
        validation_details = window_result['validation'].get('validation_details', [])
        normalized_rule_text = ' '.join(rule_text.split()).strip()
        
        # 先使用既有驗證結果
        for validation in validation_details:
            validation_rule = validation.get('rule', '')
            normalized_validation_rule = ' '.join(validation_rule.split()).strip()
            
            if normalized_validation_rule == normalized_rule_text:
                is_valid = validation.get('is_valid', False)
                status = 'existing_valid' if is_valid else 'existing_invalid'
                m_info(f"      視窗{window_result.get('window_info', {}).get('window_id', 'unknown')} 使用既有驗證結果: {status}")
                return is_valid, status
        
        # 若視窗缺必要欄位或不含等式，跳過重驗證
        if self._should_skip_revalidation(rule_text, window_result):
            m_print(f"      視窗{window_result.get('window_info', {}).get('window_id', 'unknown')} 略過重驗證")
            return False, 'skip_invalid'
        
        # 進行重驗證
        m_info(f"      視窗{window_result.get('window_info', {}).get('window_id', 'unknown')} 進行重驗證規律")
        is_valid = self._validate_rule_with_window_data(rule_text, window_result)
        status = 'revalidated_valid' if is_valid else 'revalidated_invalid'
        return is_valid, status
    
    def _should_skip_revalidation(self, rule_text: str, window_result: Dict[str, Any]) -> bool:
        """
        判斷是否應跳過重驗證。
        若規律已在此視窗被標記為無效，則跳過重驗證。
        Args:
            rule_text: 規律文字
            window_result: 視窗結果
            
        Returns:
            bool: 是否應跳過重驗證
        """
        try:
            # 若 validation_details 中已存在同規律，則視為已驗證過
            validation_details = window_result.get('validation', {}).get('validation_details', [])
            normalized_rule_text = ' '.join(rule_text.split()).strip()
            
            # iterate per-window validation details
            for validation in validation_details:
                validation_rule = validation.get('rule', '')
                normalized_validation_rule = ' '.join(validation_rule.split()).strip()
                
                # 已有同規律的驗證紀錄，跳過重驗證
                if normalized_rule_text == normalized_validation_rule:
                    m_print(f"跳過重驗證 - 已有同規律驗證紀錄: {rule_text}")
                    return True
            
            return False
            
        except Exception as e:
            m_print(f"判斷是否跳過重驗證時發生錯誤: {e}")
            return False
    
    def _are_similar_rules(self, rule1: str, rule2: str) -> bool:
        """
        判斷兩條規律是否相似（以關鍵字重疊為主）。
        Args:
            rule1: 規律1
            rule2: 規律2
            
        Returns:
            bool: 是否相似
        """
        # extract simple textual keywords
        def extract_keywords(rule):
            keywords = set()
            for word in str(rule).split():
                if any(ch.isalpha() for ch in word):
                    keywords.add(word)
            return keywords
        keywords1 = extract_keywords(rule1)
        keywords2 = extract_keywords(rule2)
        
        # 關鍵字重疊達 2 個以上即視為相似
        return len(keywords1.intersection(keywords2)) >= 2
    
    def _validate_rule_with_window_data(self, rule_text: str, window_result: Dict[str, Any]) -> bool:
        """
        使用視窗資料重驗證規律，優先採用 equation_sides。
        Args:
            rule_text: 規律文字
            window_result: 視窗結果
            
        Returns:
            bool: 是否驗證有效
        """
        try:
            # 取出視窗數值
            window_values = self._extract_window_values(window_result)
            m_info(f"        視窗數值: {window_values}")
            if not window_values or self._count_numeric_cells(window_values) < 2:
                m_print(f"        視窗數值不足，無法重驗證")
                return False

            # 嘗試找出該規律的 equation_sides
            equation_sides = self._find_equation_sides_for_rule(rule_text, window_result)

            # 若可用，將 equation_sides 附加回 validation_details（便於輸出）
            if 'validation' in window_result:
                vd = window_result['validation'].get('validation_details', [])
                for d in vd:
                    if equation_sides:
                        d['equation_sides'] = equation_sides
                        self._attach_equation_preview(d)
                    break
            
            # 直接進行數學驗證
            result = self._verify_rule_mathematically_direct(rule_text, window_values, equation_sides)
            m_info(f"        數學驗證結果: {'有效' if result else '無效'}")
            return result
            
        except Exception as e:
            m_print(f"        重驗證規律失敗: {rule_text}, 錯誤: {e}")
            return False
    
    def _verify_rule_mathematically_direct(self, rule_text: str, values: List[List[Any]], equation_sides: List[str] = None) -> bool:
        """
        直接驗證規律的數學一致性（需提供 equation_sides）。
        
        Args:
            rule_text: 規律文字（用於錯誤訊息）
            values: 數值矩陣
            equation_sides: 左右式，例如 ["$0 + $2", "$1"]
            
        Returns:
            bool: 是否驗證成功
        """
        if not values or self._count_numeric_cells(values) < 2:
            return False

        tolerance = 0.01
        
        try:
            # 必須有可用的 equation_sides 才能驗證
            if equation_sides and len(equation_sides) == 2:
                if not self._equation_sides_in_range(equation_sides, values):
                    m_info(f"          equation_sides索引超出範圍，略過驗證: {equation_sides}")
                    return False
                m_info(f"          使用 equation_sides 驗證: {equation_sides}")
                return self._verify_equation_sides(equation_sides, values, tolerance)
            else:
                # 無可用 equation_sides，無法驗證
                m_print(f"          缺少 equation_sides，無法驗證: {rule_text}")
                return False
                
        except Exception as e:
            m_print(f"          數學驗證失敗: {rule_text}, 錯誤: {e}")
            return False
    
    def _verify_equation_sides(self, equation_sides: List[str], values: List[List[Any]], tolerance: float) -> bool:
        """
        驗證 equation_sides 兩側是否近似相等。
        Args:
            equation_sides: 左右式列表，例如 ["$0 + $2", "$1"]
            values: 視窗數值矩陣
            tolerance: 容許誤差
            
        Returns:
            bool: 是否驗證成功
        """
        try:
            if len(equation_sides) != 2:
                return False
            
            left_expr, right_expr = equation_sides
            
            # 計算左側/右側表達式數值
            left_value = self._evaluate_dollar_expression(left_expr, values)
            right_value = self._evaluate_dollar_expression(right_expr, values)
            
            # 比較兩側數值
            m_info(f"equation_sides: {equation_sides}, left_value: {left_value}, right_value: {right_value}")
            return abs(left_value - right_value) < tolerance
            
        except Exception as e:
            m_print(f"解析 equation_sides 失敗: {equation_sides}, 錯誤: {e}")
            return False

    def _equation_sides_in_range(self, equation_sides: List[str], values: List[List[Any]]) -> bool:
        """
        檢查 equation_sides 中所有索引是否在 values 的範圍內。
        """
        import re

        def _to_matrix(v):
            if not isinstance(v, list):
                return []
            if v and not isinstance(v[0], list):
                return [[x] for x in v]
            return v

        matrix = _to_matrix(values)
        max_rows = len(matrix)
        max_cols = max((len(r) for r in matrix), default=0)

        def _check_expr(expr: str) -> bool:
            s = re.sub(r'(?<!\$)\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)', r'$(\1,\2)', str(expr))
            found = False
            for r_str, c_str in re.findall(r'\$\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)', s):
                found = True
                r = int(r_str)
                c = int(c_str)
                if r < 0 or c < 0 or r >= max_rows or c >= max_cols:
                    return False
            for r_str in re.findall(r'\$(\d+)', s):
                found = True
                r = int(r_str)
                if r < 0 or r >= max_rows or 0 >= max_cols:
                    return False
            return found

        if not equation_sides or len(equation_sides) != 2:
            return False
        return _check_expr(equation_sides[0]) and _check_expr(equation_sides[1])
    
    def _evaluate_dollar_expression(self, expression: str, values: List[List[Any]]) -> float:
        """
        解析 $(a,b) 索引運算式（例如 "$(0,0) + $(1,0)"）。
        Args:
            expression: 運算式，例如 "$(0,0) + $(1,0)" 或 "$(0,1)"
            values: 數值矩陣
        Returns:
            float: 計算結果
        """
        import re
        
        def _to_matrix(v):
            if not isinstance(v, list):
                return []
            if v and not isinstance(v[0], list):
                return [[x] for x in v]
            return v

        matrix = _to_matrix(values)
        max_rows = len(matrix)
        max_cols = max((len(r) for r in matrix), default=0)
        matrix = [list(r) + [None] * (max_cols - len(r)) for r in matrix]

        def _get_value(r: int, c: int) -> float:
            if r < 0 or r >= max_rows or c < 0 or c >= max_cols:
                raise ValueError(f"索引超出範圍: ({r},{c})")
            val = matrix[r][c]
            if val is None:
                raise ValueError(f"索引 ({r},{c}) 對應值為空")
            try:
                return float(val)
            except Exception:
                raise ValueError(f"索引 ({r},{c}) 對應值無法轉為數值: {val}")

        # Normalize bare "(r,c)" into "$(r,c)" before evaluation
        expression = re.sub(r'(?<!\$)\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)', r'$(\1,\2)', str(expression))

        # 先替換 $(a,b)
        def replace_2d(match):
            r = int(match.group(1))
            c = int(match.group(2))
            return str(_get_value(r, c))

        numeric_expression = re.sub(r'\$\(\s*(\d+)\s*,\s*(\d+)\s*\)', replace_2d, expression)

        # 相容舊式 $i（視為 $(i,0)）
        def replace_1d(match):
            idx = int(match.group(1))
            return str(_get_value(idx, 0))

        numeric_expression = re.sub(r'\$(\d+)', replace_1d, numeric_expression)
        
        # 安全計算數學式
        return self._safe_eval_math_expression_simple(numeric_expression)
    
    def _safe_eval_math_expression_simple(self, expression: str) -> float:
        """
        安全計算純數學運算式（僅允許數字與 +-*/()）。
        Args:
            expression: 數學運算式，例如 "237.6 + 1.2"
            
        Returns:
            float: 計算結果
        """
        import re
        
        # 僅允許安全字元
        if not re.match(r'^[\d\.\+\-\*\/\(\)\s]+$', expression):
            raise ValueError(f"運算式包含不安全字元: {expression}")
        
        # evaluate math expression in constrained context
        try:
            result = eval(expression, {"__builtins__": {}}, {})
            return float(result)
        except Exception as e:
            raise ValueError(f"無法計算運算式: {expression}, 錯誤: {str(e)}")
    
    def _find_equation_sides_for_rule(self, rule_text: str, window_result: Dict[str, Any]) -> Optional[List[str]]:
        """
        嘗試從 llm_result 中找出與規律文字對應的 equation_sides；若無則從 rule/equation 解析含 $ 索引的等式並拆成左右式。
        會比對 equation（或 rule、description）與規律文字是否一致，再取出 $i 形式的左右式。
        回傳形如 ["$0 + $1", "$2"] 的 list，找不到則為 None。
        """
        try:
            # 1) 從 llm_result.rules 尋找對應規律
            llm_rules = window_result.get('llm_result', {}).get('rules', [])
            norm = lambda s: ' '.join((s or '').split()).strip()

            for r in llm_rules:
                txt = r.get('equation', r.get('rule', r.get('description', '')))
                if norm(txt) == norm(rule_text):
                    # 直接取得 equation_sides
                    if isinstance(r.get('equation_sides'), list) and len(r['equation_sides']) == 2:
                        return r['equation_sides']
                    # 若只有 equation 含 $ 索引，拆成左右式
                    eq = r.get('equation') or r.get('equation_with_indices') or r.get('equation_with_values')
                    if isinstance(eq, str) and '$' in eq and '=' in eq:
                        left, right = eq.split('=', 1)
                        return [left.strip(), right.strip()]

            # 2) if rule text already contains a $-indexed equation, split sides directly
            if '$' in rule_text and '=' in rule_text:
                left, right = rule_text.split('=', 1)
                return [left.strip(), right.strip()]

            # map terms in rule text to row_names and build equation_sides
            row_names = window_result.get('prompt_data', {}).get('row_names') or []
            if not row_names:
                row_names = window_result.get('window_info', {}).get('index_names') or []
            mapped = self._map_terms_to_indices(rule_text, row_names)
            if mapped and len(mapped) >= 3:
                # default to A + B = C (row-based, column 0)
                return [f"$({mapped[0]},0) + $({mapped[1]},0)", f"$({mapped[2]},0)"]
            return None
        except Exception as e:
            m_print(f"_find_equation_sides_for_rule 失敗: {e}")
            return None

    def _map_terms_to_indices(self, rule_text: str, index_names: List[str]) -> Optional[List[int]]:
        """
        將規律文字中的關鍵詞（如 computed / input / target）對應到 row_names 的索引順序。
        回傳索引列表（例如 [0,1,2]），供對應 $(r,0) 形式占位符。
        """
        if not index_names:
            return None

        # keyword aliases for locating key rows/columns
        keywords = [
            ("computed_value", ["computed_value", "computed", "calc"]) ,
            ("input_value", ["input_value", "input", "raw"]) ,
            ("target_value", ["target_value", "target", "result"]) ,
        ]

        def find_index(alias_list):
            # scan index_names and find first matching alias
            for i, name in enumerate(index_names):
                name_lc = str(name).lower()
                for alias in alias_list:
                    if alias.lower() in name_lc:
                        return i
            return None

        # 依規律文字中的關鍵詞出現順序，對應到 index_names 索引（輔助關鍵詞對應，非完整解析）
        order = []
        lower_text = rule_text.lower()
        for canonical, alias_list in keywords:
            if any(a in lower_text for a in alias_list):
                idx = find_index(alias_list)
                if idx is not None:
                    order.append((lower_text.find(alias_list[0]), idx))
        if not order:
            return None

        # 依關鍵詞在規律文字中的位置排序，決定索引順序
        order.sort(key=lambda x: x[0])
        return [idx for _, idx in order]

    def _attach_equation_preview(self, detail: Dict[str, Any]) -> None:
        """
        在 validation_detail 上附加可讀的 equation 預覽字串，便於除錯與閱讀。
        範例: "$0 + $1 = $2"
        """
        eq = detail.get('equation_sides')
        if isinstance(eq, list) and len(eq) == 2:
            detail['equation_preview'] = f"{eq[0]} = {eq[1]}"

    def _extract_window_values(self, window_result: Dict[str, Any]) -> List[List[Any]]:
        """Extract numeric matrix values from a window_result structure."""
        if 'prompt_data' in window_result and 'values' in window_result['prompt_data']:
            values = window_result['prompt_data']['values']
            if isinstance(values, list):
                if values and isinstance(values[0], list):
                    return values
                return [[v] for v in values]
            return []

        if 'window_info' in window_result and 'values' in window_result['window_info']:
            values = window_result['window_info']['values']
            if isinstance(values, list) and values:
                if isinstance(values[0], list):
                    return values
                return [[v] for v in values]

        return []

    def _count_numeric_cells(self, matrix: List[List[Any]]) -> int:
        count = 0
        for row in matrix:
            for v in row:
                if isinstance(v, (int, float)):
                    count += 1
        return count
    
    def _check_rule_consistency(self, group_results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        檢查同一 start_loc_row 群組內規律的一致性。
        Args:
            group_results: 同一 start_loc_row 的視窗結果列表
        Returns:
            List[Dict]: 一致性足夠的規律列表
        """
        if len(group_results) < 2:
            # with one window only, just return its rules
            if group_results and 'llm_result' in group_results[0]:
                return group_results[0]['llm_result'].get('rules', [])
            return []
        # 統計各規律出現次數
        rule_text_patterns = defaultdict(int)
        equation_to_description = {}  # equation 對應的 description 次數
        
        for result in group_results:
            if 'llm_result' in result and 'rules' in result['llm_result']:
                for rule in result['llm_result']['rules']:
                    equation = rule.get('equation', '')
                    description = rule.get('description', '')
                    
                    if equation:
                        # 使用正規化後的 equation 作為 key
                        normalized_equation = ' '.join(equation.split()).strip()
                        rule_text_patterns[normalized_equation] += 1
                        
                        # 記錄每個 equation 對應的 description 次數
                        if normalized_equation not in equation_to_description:
                            equation_to_description[normalized_equation] = defaultdict(int)
                        equation_to_description[normalized_equation][description] += 1
        
        # keep patterns that appear often enough across windows
        min_occurrence = max(2, len(group_results) // 2)
        consistent_patterns = [pattern for pattern, count in rule_text_patterns.items() if count >= min_occurrence]
        
        # 組出符合門檻的規律，並帶入最常見的 description
        consistent_rules = []
        for result in group_results:
            if 'llm_result' in result and 'rules' in result['llm_result']:
                for rule in result['llm_result']['rules']:
                    equation = rule.get('equation', '')
                    if equation:
                        normalized_equation = ' '.join(equation.split()).strip()
                        if normalized_equation in consistent_patterns:
                            # 僅保留達標的規律，並套用最常見的 description
                            consistent_rule = rule.copy()
                            
                            # 取得此 equation 最常見的 description
                            if normalized_equation in equation_to_description:
                                most_common_description = max(
                                    equation_to_description[normalized_equation].items(),
                                    key=lambda x: x[1]
                                )[0]
                                consistent_rule['description'] = most_common_description
                            
                            consistent_rules.append(consistent_rule)
                            break  # 此視窗已納入一條符合規律，避免重複
        
        return consistent_rules
    
    def _calculate_rule_consistency_rate(self, rule_text: str, analysis_result: Dict[str, Any]) -> float:
        """
        計算單一規律在整體分析中的一致性比率。
        
        Args:
            rule_text: 規律文字
            analysis_result: 完整分析結果
        Returns:
            float: 一致性比率 (0.0 - 1.0)
        """
        # 規律文字標準化
        normalized_rule = ' '.join(rule_text.split()).strip()
        
        # 依 start_loc_row 分組統計規律出現次數
        start_loc_row_groups = defaultdict(list)
        
        for window_result in analysis_result.get('detailed_results', []):
            start_loc_row = window_result.get('window_info', {}).get('start_loc_row_name', 'unknown')
            start_loc_row_groups[start_loc_row].append(window_result)
        
        total_consistency_rate = 0.0
        valid_start_loc_rows = 0
        
        for start_loc_row, group_results in start_loc_row_groups.items():
            if len(group_results) < 2:
                continue
            # 統計此規律在該 start_loc_row 各視窗是否出現
            rule_occurrences = 0
            for result in group_results:
                if 'llm_result' in result and 'rules' in result['llm_result']:
                    for rule in result['llm_result']['rules']:
                        # 以正規化 equation 與目標規律比對
                        current_equation = rule.get('equation', '')
                        if current_equation:
                            current_normalized = ' '.join(current_equation.split()).strip()
                            if current_normalized == normalized_rule:
                                rule_occurrences += 1
                                break  
            # 計算此 start_loc_row 的一致性比率
            window_count = len(group_results)
            consistency_rate = rule_occurrences / window_count
            
            total_consistency_rate += consistency_rate
            valid_start_loc_rows += 1
        
        # 對所有有效 start_loc_row 取平均一致性比率
        return total_consistency_rate / valid_start_loc_rows if valid_start_loc_rows > 0 else 0.0
    
    def generate_rule_summary(self, analysis_result: Dict[str, Any], consistency_threshold: float = 0.8) -> Dict[str, Any]:
        """
        彙整分析結果產生規律摘要與統計（含一致性門檻篩選）。
        Args:
            analysis_result: 分析結果
            consistency_threshold: 一致性門檻，預設 80%
        Returns:
            Dict: 含規律列表與統計資訊
        """
        if not analysis_result.get('success'):
            return {'rules': [], 'statistics': {}}
        
        # 初始化各規律的統計結構
        rule_stats = defaultdict(lambda: {
            'rule_text': '',
            'total_validations': 0,
            'valid_validations': 0,
            'valid_ratio': 0.0,
            'examples': []
        })
        
        # aggregate stats from each analyzed window
        for window_result in analysis_result.get('detailed_results', []):
                rules = window_result['llm_result']['rules']
                validation = window_result.get('validation', {})
                
                for rule in rules:
                    # 以 equation 作為規律主鍵，description 供顯示
                    rule_equation = rule.get('equation', '')
                    rule_description = rule.get('description', '')
                    
                    if not rule_equation:
                        continue
                    
                    # 以 equation 作為統計用的規律識別字串
                    rule_text = rule_equation
                    
                    # description 可能為空，先正規化規律文字以便彙整
                    normalized_rule = ' '.join(rule_text.split())
                    
                    # 更新統計
                    rule_stats[normalized_rule]['rule_text'] = normalized_rule
                    rule_stats[normalized_rule]['rule_equation'] = rule_equation
                    rule_stats[normalized_rule]['rule_description'] = rule_description
                    rule_stats[normalized_rule]['total_validations'] += 1
                    
                    # 檢查此規律在該視窗驗證細節是否有效
                    is_valid = False
                    validation_details = validation.get('validation_details', validation.get('rule_validations', []))
                    
                    for rule_validation in validation_details:
                        # 以 equation 比對驗證項目
                        validation_equation = rule_validation.get('equation', rule_validation.get('rule', ''))
                        validation_normalized = ' '.join(validation_equation.split())
                        if validation_normalized == normalized_rule and rule_validation.get('is_valid', False):
                            is_valid = True
                            break
                    
                    if is_valid:
                        rule_stats[normalized_rule]['valid_validations'] += 1
                    
                    # 保留最多 3 筆範例視窗
                    if len(rule_stats[normalized_rule]['examples']) < 3:
                        example = {
                            'window_id': window_result.get('window_info', {}).get('window_id', 'unknown'),
                            'start_loc_row': window_result.get('window_info', {}).get('start_loc_row_name', 'unknown'),
                            'is_valid': is_valid
                        }
                        rule_stats[normalized_rule]['examples'].append(example)
        
        # 計算有效率
        for rule_key in rule_stats:
            total = rule_stats[rule_key]['total_validations']
            valid = rule_stats[rule_key]['valid_validations']
            rule_stats[rule_key]['valid_ratio'] = valid / total if total > 0 else 0.0
        
        # 依有效率排序
        sorted_rules = sorted(rule_stats.items(), key=lambda x: x[1]['valid_ratio'], reverse=True)
        
        # 從分析結果取出 start_loc_row 驗證細節
        start_loc_row_validation = analysis_result.get('start_loc_row_validation', {})
        start_loc_row_details = start_loc_row_validation.get('details', {})
        
        # 組出規律摘要列表（僅納入通過一致性門檻者）
        rule_summary = []
        filtered_rules = []
        
        for rule_text, stats in sorted_rules:
            if stats['total_validations'] > 0:
                # 計算此規律在 start_loc_row 的一致性
                consistency_rate = self._calculate_rule_consistency_rate(rule_text, analysis_result)
                stats['consistency_rate'] = consistency_rate
                
                # keep only rules passing consistency threshold
                if consistency_rate >= consistency_threshold:
                    filtered_rules.append((rule_text, stats))
                else:
                    # 未達門檻者記錄原因
                    stats['filtered_reason'] = (
                        f'一致性比率 {consistency_rate:.1%} 低於門檻 {consistency_threshold:.1%}'
                    )
        
        # 整體統計
        total_rules_found = len(sorted_rules)
        rules_passed_threshold = len(filtered_rules)
        
        # 彙總驗證次數統計
        total_validations = sum(stats['total_validations'] for _, stats in sorted_rules)
        text_valid_validations = sum(stats['valid_validations'] for _, stats in sorted_rules)
        consistency_valid_validations = sum(stats['valid_validations'] for _, stats in filtered_rules)
        
        statistics = {
            'total_unique_rules': rules_passed_threshold,  # 通過一致性門檻的規律數
            'total_rules_found': total_rules_found,        # 偵測到的相異規律總數
            'total_rule_validations': total_validations,   # 規則驗證總次數
            'text_valid_validations': text_valid_validations,  # 文字解析有效次數
            'consistency_valid_validations': consistency_valid_validations,  # 一致性驗證有效次數
            'consistency_threshold': consistency_threshold,
            'consistency_pass_rate': rules_passed_threshold / total_rules_found if total_rules_found > 0 else 0.0,
            'text_valid_ratio': text_valid_validations / total_validations if total_validations > 0 else 0.0,
            'consistency_valid_ratio': consistency_valid_validations / total_validations if total_validations > 0 else 0.0,
            'rule_details': {}
        }
        
        # build rule_details with consistency stats
        for rule_text, stats in sorted_rules:
                rule_detail = {
                    'validations': stats['total_validations'],
                    'valid_count': stats['valid_validations'],
                    'valid_ratio': stats['valid_ratio'],
                    'examples': stats['examples']
                }
                
                # 附加一致性統計欄位
                if 'consistency_rate' in stats:
                    rule_detail['consistency_rate'] = stats['consistency_rate']
                    rule_detail['passes_consistency_threshold'] = stats['consistency_rate'] >= consistency_threshold
                
                # 未通過篩選時記錄原因
                if 'filtered_reason' in stats:
                    rule_detail['filtered_reason'] = stats['filtered_reason']
                
                statistics['rule_details'][rule_text] = rule_detail
        
        return {
            'rules': rule_summary,
            'statistics': statistics
        }
    
    def save_rule_summary(
        self,
        analysis_result: Dict[str, Any],
        output_dir: str = None,
        consistency_threshold: float = 0.8,
        file_suffix: str = '',
        print_summary: bool = True,
    ) -> str:
        """
        將規律摘要寫入 discovered_rules.json、rule_statistics.json 等輸出檔。
        Args:
            analysis_result: 分析結果
            output_dir: 輸出目錄（可選）
            consistency_threshold: 一致性門檻，預設 80%
            file_suffix: 檔名後綴（例：_phase1、_phase2_final）
            print_summary: 是否列印規律摘要到 console
        Returns:
            str: 主要規律 JSON 檔路徑
        """
        if not analysis_result.get('success'):
            m_print("分析結果不成功，無法產生規律摘要")
            return None
        
        # 產生規律摘要（套用一致性門檻篩選）
        summary = self.generate_rule_summary(analysis_result, consistency_threshold)
        
        # 預設輸出路徑
        if output_dir is None:
            outputs_dir = Path(__file__).parent.parent / 'outputs'
            outputs_dir.mkdir(exist_ok=True)
            output_dir = outputs_dir
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(exist_ok=True)
        
        # 寫入 discovered_rules.json（可附加檔名後綴）
        rules_file = _with_file_suffix(output_dir / 'discovered_rules.json', file_suffix)
        try:
            # 透過 LOGger.dcp 正規化門檻值（相容預設 0.8 等行為）
            threshold = LOGger.dcp(consistency_threshold)

            # overall_validation 來自 _retrospective_validate_start_loc_rows(...)
            overall_validation = analysis_result.get('start_loc_row_validation', {})
            details_by_start_loc = overall_validation.get('details', {})

            by_start_loc_row = {}
            start_loc_rows_with_any_pass = 0

            for start_loc, group in details_by_start_loc.items():
                # 整理每個 start_loc_row 的通過規律
                start_loc_row_name = group.get('start_loc_row_name', 'unknown')
                start_loc_row_indicated = group.get('start_loc_row_indicated', [])
                rule_details = group.get('rule_details', {})
                window_count = group.get('window_count', 0)

                # 必須使用回溯驗證階段已套用 degeneracy 過濾後的 passed_rules；
                # 若僅從 rule_details 依門檻重掃，會把已剔除的 trivial 規律又寫回 discovered_rules.json。
                filtered_map = group.get('passed_rules')
                if isinstance(filtered_map, dict):
                    source_items = list(filtered_map.items())
                else:
                    source_items = [
                        (k, v)
                        for k, v in rule_details.items()
                        if v.get('consistency_rate', 0.0) >= threshold
                    ]

                passed_rules = []
                for rule_text, r in source_items:
                    cr = r.get('consistency_rate', 0.0)
                    passed_rules.append({
                        "rule": rule_text,
                        "description": r.get("rule_description", rule_text),
                        "equation": r.get("rule_equation", rule_text),
                        "equation_sides": r.get("equation_sides"),
                        "valid_windows": r.get("valid_windows", 0),
                        "total_windows": r.get("total_windows", window_count),
                        "consistency_rate": cr,
                        "supporting_windows": r.get("validation_details", []),
                    })

                if passed_rules:
                    start_loc_rows_with_any_pass += 1

                block = {
                    "start_loc_row_name": start_loc_row_name,
                    "start_loc_row_indicated": start_loc_row_indicated,
                    "window_count": window_count,
                    "total_rules_found": len(rule_details),
                    "passed_rules_count": len(passed_rules),
                    "passed_rules": passed_rules,
                }

                by_start_loc_row[start_loc] = block
            
            with open(rules_file, 'w', encoding='utf-8') as f:
                json.dump(by_start_loc_row, f, ensure_ascii=False, indent=4)

        except Exception as e:
            m_print(f"儲存規律 JSON 失敗: {e}")
            return None


        
        # 儲存統計檔案
        stats_file = _with_file_suffix(output_dir / 'rule_statistics.json', file_suffix)
        try:
            stats_data = {
                'analysis_info': {
                    'excel_file': analysis_result['excel_file'],
                    'analysis_time': analysis_result['analysis_time'],
                    'window_shape': analysis_result['window_scanning']['window_shape'],
                    'total_windows': analysis_result['window_scanning']['total_windows'],
                    'analyzed_windows': analysis_result['llm_analysis']['analyzed_windows'],
                    'llm_mode': analysis_result['llm_analysis']['llm_mode']
                },
                'summary': {
                    'total_rules_found': summary['statistics']['total_rules_found'],
                    'total_unique_rules': summary['statistics']['total_unique_rules'],
                    'total_validations': summary['statistics']['total_rule_validations'],
                    'text_valid_validations': summary['statistics']['text_valid_validations'],
                    'consistency_valid_validations': summary['statistics']['consistency_valid_validations'],
                    'consistency_threshold': summary['statistics']['consistency_threshold'],
                    'consistency_pass_rate': summary['statistics']['consistency_pass_rate'],
                    'text_valid_ratio': summary['statistics']['text_valid_ratio'],
                    'consistency_valid_ratio': summary['statistics']['consistency_valid_ratio']
                },
                'start_loc_row_validation': analysis_result['start_loc_row_validation'],  # 起始列驗證結果
                'rule_details': summary['statistics']['rule_details']
            }
            
            with open(stats_file, 'w', encoding='utf-8') as f:
                json.dump(stats_data, f, ensure_ascii=False, indent=4)
            m_print(f"已儲存規律統計檔案: {stats_file}")
        except Exception as e:
            m_print(f"儲存規律統計時發生錯誤: {e}")
        
        # 顯示摘要統計
        if print_summary:
            print(f"\n=== 規律摘要（含一致性門檻 {consistency_threshold:.1%}）===")
            print(f"偵測到規律總數: {summary['statistics']['total_rules_found']}")
            print(f"通過一致性門檻規律數: {summary['statistics']['total_unique_rules']}")
            print(f"一致性門檻通過率: {summary['statistics']['consistency_pass_rate']:.1%}")
            print(f"\n=== 驗證統計 ===")
            print(f"規則驗證總次數: {summary['statistics']['total_rule_validations']}")
            print(f"文字解析有效次數: {summary['statistics']['text_valid_validations']} ({summary['statistics']['text_valid_ratio']:.1%})")
            print(f"一致性驗證有效次數: {summary['statistics']['consistency_valid_validations']} ({summary['statistics']['consistency_valid_ratio']:.1%})")

            print(f"\n=== 通過一致性門檻規律列表 ===")
            for i, rule in enumerate(summary['rules'][:10], 1):
                stats = summary['statistics']['rule_details'][rule]
                consistency_rate = stats.get('consistency_rate', 0.0)
                print(f"{i}. {rule}")
                print(f"   驗證次數: {stats['validations']} 次, 有效次數: {stats['valid_count']} 次, 有效率: {stats['valid_ratio']:.1%}")
                print(f"   一致性比率: {consistency_rate:.1%}")
            
            if len(summary['rules']) > 10:
                print(f"... and {len(summary['rules']) - 10} more rules")
        
        return str(rules_file)
    
    def save_analysis_snapshot(
        self,
        analysis_result: Dict[str, Any],
        output_file: str = None,
        file_suffix: str = '',
    ) -> str:
        """將完整 analysis_result 寫成快照 JSON（可用於 phase1/phase2 分段產物）。"""
        if not analysis_result.get('success'):
            m_print("analysis_result is not successful, skip saving analysis snapshot")
            return None

        if output_file is None:
            outputs_dir = Path(__file__).parent.parent / 'outputs'
            outputs_dir.mkdir(exist_ok=True)
            output_path = _with_file_suffix(outputs_dir / 'analysis_result.json', file_suffix)
        else:
            output_path = _with_file_suffix(Path(output_file), file_suffix)

        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(analysis_result, f, ensure_ascii=False, indent=2, default=str)
            m_print(f"analysis snapshot saved: {output_path}")
            return str(output_path)
        except Exception as e:
            m_print(f"save analysis snapshot failed: {e}")
            return None

    def save_to_observed_json(
        self,
        analysis_result: Dict[str, Any],
        output_file: str = None,
        file_suffix: str = '',
    ) -> str:
        """Save analysis output into observed.json-style structure."""
        if not analysis_result.get('success'):
            m_print("analysis_result is not successful, skip saving observed json")
            return None

        if output_file is None:
            outputs_dir = Path(__file__).parent.parent / 'outputs'
            outputs_dir.mkdir(exist_ok=True)
            output_path = _with_file_suffix(outputs_dir / 'observed.json', file_suffix)
        else:
            output_path = _with_file_suffix(Path(output_file), file_suffix)

        observed_entry = {
            'execution_time': analysis_result.get('analysis_time'),
            'prompt_template': 'window-based rule analysis prompt',
            'observed_file_path': analysis_result.get('excel_file'),
            'observed_sheet': None,
            'observed_features': (
                f"window_shape={self.window_shape}, "
                f"start_loc_row_name={analysis_result.get('window_scanning', {}).get('start_loc_row_filter')}"
            ),
            'total_data_points': analysis_result.get('llm_analysis', {}).get('analyzed_windows'),
            'valid_rules_count': analysis_result.get('llm_analysis', {}).get('valid_rules_found'),
            'success_rate': analysis_result.get('llm_analysis', {}).get('overall_success_rate'),
            'llm_mode': analysis_result.get('llm_analysis', {}).get('llm_mode'),
            'start_loc_row_validation': analysis_result.get('start_loc_row_validation'),
            'detailed_analysis': analysis_result.get('detailed_results', []),
        }

        observed_data = {str(Path(__file__).resolve()): observed_entry}

        if output_path.exists():
            try:
                with open(output_path, 'r', encoding='utf-8') as f:
                    existing_data = json.load(f)
                if isinstance(existing_data, dict):
                    existing_data.update(observed_data)
                    observed_data = existing_data
            except Exception as e:
                m_print(f"load existing observed json failed: {e}")

        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(observed_data, f, ensure_ascii=False, indent=2, default=str)
            m_print(f"observed json saved: {output_path}")
            return str(output_path)
        except Exception as e:
            m_print(f"save observed json failed: {e}")
            return None

    def save_timer(self, output_file: str = None, file_suffix: str = '') -> bool:
        """
        將各階段 timer 資訊寫入 timer.json。
        """
        self.timer['partial_seconds'] = {}
        self.timer['accumulated_seconds'] = {}
        if output_file is None:
            outputs_dir = Path(__file__).parent.parent / "outputs"
            outputs_dir.mkdir(exist_ok=True)
            output_file = _with_file_suffix(outputs_dir / "timer.json", file_suffix)
        else:
            output_file = _with_file_suffix(Path(output_file), file_suffix)
        # 將各階段時間戳排序後計算區段耗時
        # timings = np.array(tuple(zip(**tuple(self.timer.items()))))
        timings = np.array(
            sorted(
                [
                    (k, v.timestamp())
                    for k, v in self.timer.items()
                    if k not in ("partial_seconds", "accumulated_seconds")
                    and hasattr(v, "timestamp")
                ],
                key=lambda x: x[1]
            ),
            dtype=object
        )
        if timings.size == 0:
            m_print("timer 尚無可用時間戳，略過儲存")
            return False
        t0 = np.min(timings[:, 1])
        order = np.argsort(timings[:, 1])   # 依時間欄位排序
        timings = timings[order]
        ti = LOGger.dcp(t0)
        for i,tim in enumerate(timings[1:, 1]):
            # self.timer['partial_seconds'][f"{timings[i-1, 0]} to {timings[i, 0]}"] = float((tim - ti).total_seconds())
            # self.timer['accumulated_seconds'][f"{timings[0, 0]} to {timings[i, 0]}"] = float((tim - t0).total_seconds())
            self.timer['partial_seconds'][f"{timings[i-1, 0]} to {timings[i, 0]}"] = float(tim - ti)
            self.timer['accumulated_seconds'][f"{timings[0, 0]} to {timings[i, 0]}"] = float(tim - t0)
            ti = LOGger.dcp(tim)
            
        LOGger.CreateFile(output_file, lambda f: LOGger.save_json(self.timer, f, indent=4))
        return True

def build_arg_parser():
    parser = argparse.ArgumentParser(description='Math rule analysis runner')
    parser.add_argument('excel_file', nargs='?', default='correct_simple.xlsx',
                        help='Excel file path (default: correct_simple.xlsx)')
    parser.add_argument('--use-openai', action='store_true',
                        help='Use TextProcessor OpenAI provider instead of remote8b')
    parser.add_argument('--openai-model', type=str, default='gpt35_chat',
                        help='OpenAI model alias (default: gpt35_chat)')
    parser.add_argument('-wh', '--window-height', type=int, default=3,
                        help='Window height (default: 3)')
    parser.add_argument('-ww', '--window-width', type=int, default=1,
                        help='Window width (default: 1)')
    parser.add_argument('-sr', '--step-row', type=int, default=1,
                        help='Window scan row step (default: 1)')
    parser.add_argument('-sc', '--step-col', type=int, default=1,
                        help='Window scan col step (default: 1)')
    parser.add_argument('-r', '--start-loc-row-name', type=str,
                        help='Optional start_loc_row_name filter')
    parser.add_argument('--output', type=str,
                        help='Observed output file path')
    parser.add_argument('--save-details', action='store_true',
                        help='Save detailed analysis JSON')
    parser.add_argument('--consistency-threshold', type=float, default=0.8,
                        help='Consistency threshold (default: 0.8)')
    parser.add_argument('--quick-scan-threshold', type=int, default=3,
                        help='A線：每 start_loc 群 Phase1 抽樣這麼多個視窗做 LLM，其餘延至 Phase2（0=關閉）')
    parser.add_argument('--quick-scan-seed', type=int, default=None,
                        help='A線可選：隨機洗牌種子，便於重現')
    parser.add_argument('--phase1-global-llm-batch', dest='phase1_global_llm_batch', action='store_true', default=True,
                        help='A線 Phase1：跨群湊滿 batch（global scheduler，預設）')
    parser.add_argument('--phase1-per-group-llm-batch', dest='phase1_global_llm_batch', action='store_false',
                        help='A線 Phase1：改回逐群 batch（legacy）')
    parser.add_argument('--no-phase2-overlap', action='store_true',
                        help='A線：關閉 Phase2 與 Phase1 回溯並行（改為 checkpoint 後再同步跑 Phase2）')

    return parser

def main(args):
    """CLI entrypoint."""
    window_shape = (args.window_height, args.window_width)
    step_size = (args.step_row, args.step_col)
    analyzer = MathRuleAnalyzer(
        use_openai=args.use_openai,
        window_shape=window_shape,
        step_size=step_size,
        openai_model=args.openai_model,
        consistency_threshold=args.consistency_threshold,
        quick_scan_threshold=args.quick_scan_threshold,
        quick_scan_seed=args.quick_scan_seed,
        use_phase1_global_llm_batch=args.phase1_global_llm_batch,
        phase2_overlap_phase1_retro=not args.no_phase2_overlap,
    )

    result = analyzer.analyze_excel_file(args.excel_file, args.start_loc_row_name)

    if result.get('success'):
        rules_file = analyzer.save_rule_summary(result, consistency_threshold=args.consistency_threshold)
        observed_file = analyzer.save_to_observed_json(result, args.output)
        analyzer.save_analysis_snapshot(result)

        if args.save_details:
            details_file = Path(__file__).parent / f"analysis_details_{dt.now().strftime('%Y%m%d_%H%M%S')}.json"
            try:
                with open(details_file, 'w', encoding='utf-8') as f:
                    json.dump(result, f, ensure_ascii=False, indent=2, default=str)
                m_print(f"details saved: {details_file}")
            except Exception as e:
                m_print(f"save details failed: {e}")

        analyzer.save_timer()

        # A線啟用時，另外保留一份最終結論後綴檔，避免覆蓋 Phase1 checkpoint。
        if args.quick_scan_threshold > 0:
            analyzer.save_rule_summary(
                result,
                consistency_threshold=args.consistency_threshold,
                file_suffix='_phase2_final',
                print_summary=False,
            )
            analyzer.save_to_observed_json(result, file_suffix='_phase2_final')
            analyzer.save_analysis_snapshot(result, file_suffix='_phase2_final')
            analyzer.save_timer(file_suffix='_phase2_final')
            m_print("A線 Phase2 最終結論已輸出（_phase2_final 後綴）")
        print("\n=== Analysis Summary ===")
        print(f"file: {result.get('excel_file')}")
        print(f"time: {result.get('analysis_time')}")
        print(f"analyzed_windows: {result.get('llm_analysis', {}).get('analyzed_windows')}")
        print(f"rules_found: {result.get('llm_analysis', {}).get('total_rules_found')}")
        print(f"rules_valid: {result.get('llm_analysis', {}).get('valid_rules_found')}")
        print(f"success_rate: {result.get('llm_analysis', {}).get('overall_success_rate', 0):.2%}")
        print(f"llm_mode: {result.get('llm_analysis', {}).get('llm_mode')}")
        start_loc_row_validation = result.get('start_loc_row_validation', {}) or {}
        if start_loc_row_validation:
            print("\n=== Row Group Rule Validation Summary ===")
            print(f"total_groups: {start_loc_row_validation.get('total_start_loc_rows')}")
            print(f"valid_groups: {start_loc_row_validation.get('valid_start_loc_rows')}")
            print(f"group_success_rate: {start_loc_row_validation.get('start_loc_row_success_rate', 0):.2%}")
            details = start_loc_row_validation.get('details', {}) or {}
            if details:
                print("\nper_group:")
                def _sort_key(k):
                    try:
                        return int(k)
                    except Exception:
                        return str(k)
                for start_loc in sorted(details.keys(), key=_sort_key):
                    group = details.get(start_loc, {}) or {}
                    row_name = group.get('start_loc_row_name', 'unknown')
                    window_count = group.get('window_count', 0)
                    total_rules = group.get('total_rules_found', 0)
                    passed_rules = group.get('passed_rules_count', 0)
                    passed_rate = group.get('passed_rules_rate', 0.0)
                    print(
                        f"- start_loc={start_loc} row_name={row_name} "
                        f"windows={window_count} rules={total_rules} "
                        f"passed={passed_rules} passed_rate={passed_rate:.2%}"
                    )
        def _get_ts(key: str):
            val = analyzer.timer.get(key)
            if hasattr(val, "timestamp"):
                return val.timestamp()
            if isinstance(val, (int, float)) and val > 0:
                return float(val)
            return None

        analysis_start = _get_ts('analysis_start')
        analysis_end = _get_ts('analysis_end') or dt.now().timestamp()
        table_end = _get_ts('table_detection')
        df_end = _get_ts('dataframe_extraction')
        scan_end = _get_ts('window_scanning')
        llm_end = _get_ts('llm_analysis')
        retro_end = _get_ts('retrospective_validation')

        print("\n=== Timing Summary ===")
        if analysis_start is not None:
            total_seconds = max(0.0, analysis_end - analysis_start)
            print(f"total_seconds: {total_seconds:.2f}")
        if analysis_start is not None and table_end is not None:
            print(f"table_detection_seconds: {max(0.0, table_end - analysis_start):.2f}")
        if table_end is not None and df_end is not None:
            print(f"dataframe_extraction_seconds: {max(0.0, df_end - table_end):.2f}")
        if df_end is not None and scan_end is not None:
            print(f"window_scanning_seconds: {max(0.0, scan_end - df_end):.2f}")
        if scan_end is not None and llm_end is not None:
            print(f"llm_analysis_seconds: {max(0.0, llm_end - scan_end):.2f}")
        if llm_end is not None and retro_end is not None:
            print(f"retrospective_validation_seconds: {max(0.0, retro_end - llm_end):.2f}")
        pt = result.get('llm_analysis', {}).get('pipeline_timing') or {}
        if pt.get('phase2_seconds') is not None:
            print(f"phase2_seconds: {float(pt['phase2_seconds']):.2f}")
        if pt.get('overlap_seconds') is not None:
            print(f"overlap_seconds: {float(pt['overlap_seconds']):.2f}")
        if observed_file:
            print(f"observed_file: {observed_file}")
        if rules_file:
            print(f"rules_file: {rules_file}")
            print(f"rule_stats_file: {Path(rules_file).parent / 'rule_statistics.json'}")
    else:
        print(f"analysis failed: {result.get('error', 'unknown error')}")


if __name__ == "__main__":
    parser = build_arg_parser()
    args = parser.parse_args()
    main(args)

