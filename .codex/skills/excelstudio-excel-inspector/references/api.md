# ExcelStudio Excel Inspector API

## Endpoint

- `POST http://10.1.3.127:6330/api/excel/inspect-json`

## Config

API host/port 來源：

- [config.json](C:\ML_HOME\ExcelStudio\config.json)
- 欄位：
  - `api_host`: 預設 `10.1.3.127`
  - `api_port`: 預設 `6330`

`api_server.py` 啟動時也支援環境變數覆蓋：

- `EXCELSTUDIO_API_HOST`
- `EXCELSTUDIO_API_PORT`

## Request Body

Body 為 JSON，至少需提供其中之一：

- `path`
- `file_path`
- `xlsx_path`

其他可選欄位：

- `relation_max_k`（int，預設 6）：關係推導時最多使用幾個欄位做組合嘗試

最小範例：

```powershell
curl -X POST "http://10.1.3.127:6330/api/excel/inspect-json" `
  -H "Content-Type: application/json" `
  -d '{ "path": "C:\\ML_HOME\\ExcelStudio\\defect_simple.xlsx", "relation_max_k": 6 }'
```

## Response（主要欄位）

回傳為一個 JSON（與 CLI `inspect_workbook.py` 產出的結構一致），重點欄位：

- `version`: `excel_inspector_v1`
- `generated_at`: 產生時間（ISO 字串）
- `workbook`
  - `filename`
  - `path`
  - `sheet_count`
- `sheets[]`
  - `sheet`: 工作表名
  - `shape`: `{ rows, cols }`
  - `tables[]`: 偵測到的表格區域（例如 `range: A1:L39`，以及 density/confidence）
  - `layouts[]`: header/index 推斷結果
  - `type_profiles[]`: 型別輪廓與 `anomalies`（例如期待 number 但遇到 string）
  - `relations[]`: 關係式字串（例如 `金額 = 數量 * 單價`）
  - `semantic_rule_candidates[]`: 由 relations 派生的語意規則候選
- `summary`
  - `total_sheets`
  - `total_tables`
  - `total_relations`
  - `total_type_anomalies`
- `discovered_rules`（若啟用規則輸出）
  - `version`: `excel_audit_rule_v1`
  - `workbook`: workbook 基本資訊
  - `rules[]`: 規則列表（常見 type：`table_region`、`header`、`index`、`data_type`、`arithmetic_relation`）

## MCP 注意事項

`mcp_server.py` 目前**沒有**包裝 `inspect-json` 的 MCP tool。若要透過 MCP 使用，請：

1. 確認 `api_server.py` 已啟動（預設 `config.json` 為 `10.1.3.127:6330`）
2. 直接以 HTTP 呼叫 `POST /api/excel/inspect-json`，或在本機用 CLI

## Error Codes（常見）

- 400: 缺少 `path/file_path/xlsx_path`
- 404: 找不到檔案路徑
- 500: 伺服器端 inspect 失敗（訊息中會帶 `inspect failed: ...`）

