from __future__ import annotations

from collections import deque
from typing import Any, Dict, List

import numpy as np
import pandas as pd

from .utils import excel_range, is_empty_value


def _non_empty_mask(df: pd.DataFrame) -> np.ndarray:
    return np.array([[not is_empty_value(v) for v in row] for row in df.to_numpy(dtype=object)], dtype=bool)


def detect_tables_from_sheet(
    sheet_name: str,
    df: pd.DataFrame,
    *,
    min_non_empty: int = 4,
    min_density: float = 0.12,
    bridge_gap: int = 1,
) -> List[Dict[str, Any]]:
    """Detect rectangular candidate table regions from non-empty cell components.

    bridge_gap fills very small blank gaps so semi-rectangular Excel tables are not split too aggressively.
    """
    if df.empty:
        return []
    mask = _non_empty_mask(df)
    if bridge_gap > 0:
        mask = _bridge_small_gaps(mask, bridge_gap)

    rows, cols = mask.shape
    seen = np.zeros_like(mask, dtype=bool)
    components = []
    for r in range(rows):
        for c in range(cols):
            if not mask[r, c] or seen[r, c]:
                continue
            q = deque([(r, c)])
            seen[r, c] = True
            cells = []
            while q:
                rr, cc = q.popleft()
                cells.append((rr, cc))
                for dr, dc in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                    nr, nc = rr + dr, cc + dc
                    if 0 <= nr < rows and 0 <= nc < cols and mask[nr, nc] and not seen[nr, nc]:
                        seen[nr, nc] = True
                        q.append((nr, nc))
            components.append(cells)

    tables = []
    tid = 1
    original_mask = _non_empty_mask(df)
    for cells in components:
        rs = [x[0] for x in cells]
        cs = [x[1] for x in cells]
        top, bottom = min(rs), max(rs)
        left, right = min(cs), max(cs)
        bbox_area = (bottom - top + 1) * (right - left + 1)
        non_empty = int(original_mask[top : bottom + 1, left : right + 1].sum())
        density = non_empty / max(1, bbox_area)
        if non_empty < min_non_empty or density < min_density:
            continue
        height = bottom - top + 1
        width = right - left + 1
        shape_score = min(1.0, density * 2.0) * min(1.0, non_empty / 20.0)
        tables.append(
            {
                "table_id": f"{sheet_name}_T{tid:03d}",
                "sheet": sheet_name,
                "range": excel_range(top + 1, left + 1, bottom + 1, right + 1),
                "top_row": top + 1,
                "left_col": left + 1,
                "bottom_row": bottom + 1,
                "right_col": right + 1,
                "height": height,
                "width": width,
                "non_empty_count": non_empty,
                "density": round(density, 4),
                "confidence": round(max(0.05, min(0.99, shape_score)), 4),
            }
        )
        tid += 1
    tables.sort(key=lambda x: (x["top_row"], x["left_col"]))
    return tables


def _bridge_small_gaps(mask: np.ndarray, gap: int) -> np.ndarray:
    out = mask.copy()
    rows, cols = mask.shape
    for r in range(rows):
        filled = np.where(mask[r])[0]
        if len(filled) >= 2:
            for a, b in zip(filled[:-1], filled[1:]):
                if 1 <= b - a - 1 <= gap:
                    out[r, a : b + 1] = True
    for c in range(cols):
        filled = np.where(mask[:, c])[0]
        if len(filled) >= 2:
            for a, b in zip(filled[:-1], filled[1:]):
                if 1 <= b - a - 1 <= gap:
                    out[a : b + 1, c] = True
    return out
