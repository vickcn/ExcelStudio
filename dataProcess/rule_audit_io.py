# -*- coding: utf-8 -*-
"""
規律審查工具：JSON / 文字報告讀寫。

職責：載入 observed.json、discovered_rules.json、rule_statistics.json、analysis_details_*.json；
輸出 JSON 與可選人類可讀摘要（md/txt）。
"""

from pathlib import Path
from typing import Any, Optional, Union

import json

from rule_audit_types import JsonValue


def load_json(path: Union[str, Path], encoding: str = "utf-8") -> JsonValue:
    """
    從 path 讀取 JSON。

    Returns:
        解析後的 dict 或 list（依檔案內容而定）。
    """
    p = Path(path)
    with p.open(encoding=encoding) as f:
        return json.load(f)


def write_report(
    report: Any,
    out_path: Union[str, Path],
    *,
    encoding: str = "utf-8",
    indent: int = 2,
) -> None:
    """
    將 report 寫成 JSON（開發計畫 6.1 機器可讀主輸出）。

    Args:
        report: 可 json 序列化的結構。
        out_path: 輸出檔路徑；目錄不存在時應建立（實作細節）。
    """
    p = Path(out_path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding=encoding) as f:
        json.dump(report, f, ensure_ascii=False, indent=indent, default=str)


def write_text_report(content: str, out_path: Union[str, Path], encoding: str = "utf-8") -> None:
    """
    寫入人類可讀摘要（md 或 txt，開發計畫 6.2）。

    Args:
        content: 完整文字內容。
        out_path: 輸出檔路徑。
    """
    p = Path(out_path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(content, encoding=encoding)


def load_optional_json(path: Optional[Union[str, Path]], encoding: str = "utf-8") -> Optional[JsonValue]:
    """
    可選載入；path 為 None 時回傳 None（CLI 選填參數用）。
    """
    if path is None:
        return None
    return load_json(path, encoding=encoding)
