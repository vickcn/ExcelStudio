---
name: excelstudio-excel-inspector
description: ExcelStudio 的 excel_inspector 工作簿檢查流程。用於執行 dataProcess.excel_inspector.inspect_workbook 產生 JSON 報告、解讀 summary/relations/type anomalies、以及處理常見警告與中文編碼顯示問題。
---

# ExcelStudio Excel Inspector

## 目標

- 在 `C:\ML_HOME\ExcelStudio` 內，用 CLI 或 API 檢查 `.xlsx/.xlsm/.xls/.csv`，輸出 `*.inspect.json`
- 快速確認輸出是否成功、結構是否合理，並解讀重點欄位

## 執行環境

```powershell
conda activate tfv2
cd C:\ML_HOME\ExcelStudio
```

## 呼叫方式選擇

| 方式 | 適用情境 | 備註 |
|------|----------|------|
| CLI | 本機批次、產出 `*.inspect.json` | 見下方快速使用 |
| API `POST /api/excel/inspect-json` | 遠端或 agent HTTP 呼叫 | 預設 `10.1.3.127:6330`，詳見 `references/api.md` |
| MCP (`mcp_server.py`) | Cursor / MCP Gateway | **目前沒有** `inspect-json` tool；請改走 API 或 CLI |

## 與 rules_discover / audit 的差異

- `inspect-json` / `inspect_workbook`：結構分析（table、header/index、型別、數值關係），輸出 `discovered_rules` 為輕量規則摘要。
- `rules_discover` / `audit_excel` / `full_flow`：另一條規則發現與稽核管線，產出 `outputs/discovered_rules*.json` 等檔案。
- 兩者用途不同，不要混用；若要深入規則發現與稽核，請用 MCP 的 `rules_discover`、`audit_excel`。

## API 快速使用

本 repo 也提供 API：`POST /api/excel/inspect-json`。預設 host/port 由 `config.json` 的 `api_host` / `api_port` 決定（目前為 `10.1.3.127:6330`）。

最小 curl 範例：

```powershell
curl -X POST "http://10.1.3.127:6330/api/excel/inspect-json" `
  -H "Content-Type: application/json" `
  -d '{ "path": "C:\\ML_HOME\\ExcelStudio\\defect_simple.xlsx", "relation_max_k": 6 }'
```

回傳欄位與更多說明請看 `references/api.md`。

## 快速使用 (PowerShell)

在 repo 根目錄 `C:\ML_HOME\ExcelStudio`：

```powershell
python -m dataProcess.excel_inspector.inspect_workbook .\defect_simple.xlsx `
  --relation-max-k 6 `
  --out .\inspect_reports_outputs\defect_simple.inspect.json
```

- `--relation-max-k`（預設 6）：關係推導時最多使用幾個欄位做組合嘗試；API 對應欄位 `relation_max_k`（範圍 1–20）。

## 成功與否的最小檢查

```powershell
$p = '.\\inspect_reports_outputs\\defect_simple.inspect.json'
Test-Path $p
(Get-Item $p).Length
python -c \"import json; json.load(open(r'inspect_reports_outputs\\\\defect_simple.inspect.json','r',encoding='utf-8')); print('json_ok')\"
```

## 讀中文時避免亂碼 (Windows PowerShell)

`Get-Content` 請指定 `-Encoding utf8`：

```powershell
Get-Content .\\inspect_reports_outputs\\defect_simple.inspect.json -TotalCount 80 -Encoding utf8
```

## 報告重點欄位

- `version`: `excel_inspector_v1`
- `summary.total_sheets / total_tables / total_relations / total_type_anomalies`
- `sheets[].tables[]`: 偵測到的 table range（例：`A1:L39`）、density、confidence
- `sheets[].layouts[]`: header / index 推斷結果
- `sheets[].type_profiles[].anomalies`: 型別異常（例如期望 number 但遇到 string）
- `sheets[].relations[]`: 推得的數值關係式（例：`金額 = 數量 * 單價`）
- `sheets[].semantic_rule_candidates[]`: 由 relations 派生的語意規則候選
- `discovered_rules.rules[]`: 輕量規則輸出，常見 type：
  - `table_region`、`header`、`index`、`data_type`
  - `arithmetic_relation`（confidence >= 0.85 的關係式）

## xlsx_parser 整合

`dataProcess.ContextParser.xlsx_parser` 可透過 `table_detector='inspector'`（或 `feature_scan=True`）重用 inspector 的 table detector。參數與 fallback 行為見 `references/parser-integration.md`。

## 常見警告的判讀

- `RuntimeWarning: ... found in sys.modules ... unpredictable behaviour`
  - 通常不影響輸出是否產生；若遇到「同一檔案多跑幾次結果不一致」才需要深追 import 路徑或執行方式。
- `RuntimeWarning: Mean of empty slice`
  - 多半表示某些候選關係在該 table 區間缺乏可比較樣本（空或全 NaN），通常不會導致整體檢查失敗。
