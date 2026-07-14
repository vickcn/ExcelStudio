# xlsx_parser 與 excel_inspector 整合

## 觸發方式

在 `dataProcess.ContextParser.xlsx_parser` 中，下列設定會改用 `excel_inspector.table_detector.detect_tables_from_sheet`：

- `table_detector='inspector'` 或 `table_detector='excel_inspector'`
- `feature_scan=True`（為 inspector 的 alias）

預設為 `native`（xlsx_parser 內建 detector）。

## 可調參數

透過 `inspector_table_params` 傳入：

| 參數 | 預設 | 說明 |
|------|------|------|
| `min_non_empty` / `min_non_empty_cells` | 4 | 表格候選區最少非空儲存格數 |
| `min_density` | 0.12 | 非空密度下限 |
| `bridge_gap` / `max_gap_cols` / `max_gap_rows` | 1 | 填補小空白間隙，避免表格被切太碎 |

## Fallback

若 `excel_inspector` 模組無法載入，會記錄 warning 並 fallback 至 native table detector，不會中斷整體 extract 流程。

## 與 inspect_workbook 的差異

- `inspect_workbook`：完整工作簿分析（table + layout + type + relation + discovered_rules）。
- xlsx_parser inspector 模式：僅重用 **table 偵測** 這一段，供 chunk / extract 流程選更好的表格邊界。
