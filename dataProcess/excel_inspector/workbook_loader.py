from __future__ import annotations

from pathlib import Path
from typing import Any, Dict

import pandas as pd


def load_workbook_frames(path: str | Path) -> Dict[str, pd.DataFrame]:
    """Load workbook sheets as headerless DataFrames.

    Values are kept as objects so the inspector can infer table/header/data areas itself.
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(str(p))

    suffix = p.suffix.lower()
    if suffix == ".csv":
        return {p.stem or "csv": pd.read_csv(p, header=None, dtype=object, keep_default_na=False)}
    if suffix not in {".xlsx", ".xls", ".xlsm"}:
        raise ValueError(f"Unsupported file extension: {suffix}")

    engine = "openpyxl" if suffix in {".xlsx", ".xlsm"} else None
    raw = pd.read_excel(p, sheet_name=None, header=None, dtype=object, engine=engine, keep_default_na=False)
    # Drop fully empty trailing rows/cols only; keep internal blanks.
    cleaned: Dict[str, pd.DataFrame] = {}
    for name, df in raw.items():
        df = df.copy()
        df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
        cleaned[str(name)] = df.reset_index(drop=True)
    return cleaned
