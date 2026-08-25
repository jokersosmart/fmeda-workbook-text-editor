# Feature Specification: FMEDA Workbook Text Workspace

**Feature Branch**: `fmeda-workbook-text`  
**Status**: Draft / Vertical Slice 1  
**Input**: 將原始 FMEDA Excel 轉成可追溯的純文字模型，並與 Markdown Block Editor 協同使用。

## 1. 目標與邊界

本功能的第一個 vertical slice 必須保護原始 Excel 不被修改，建立一份衍生 Excel 工作版本，並輸出可供審查者、主管、跨部門與 Editor 使用的結構化 workspace。第一版先完成單一 FMEDA 的完整匯入與驗證；未來其他類似 Excel 透過 profile detection 與 mapping 規則加入，不將第一份檔案的列號與欄位位置寫死在共通核心。

本階段不承諾脫離 Excel 的獨立重算，也不允許普通 Markdown 編輯直接覆蓋公式。Excel 的快取值必須明確標記為來源快取值，而不是宣稱為最新重算結果。

## 2. 使用者與主要價值

主要使用者為 FMEDA 工程師以外的審查者、主管與跨部門人員；FMEDA 工程師則需要能由結果回查公式與來源儲存格。Markdown Block Editor 是協同閱讀、審查、註記與受控編輯入口，不取代 Excel 的既有專業工作流程。

## 3. User Stories

### US-001：保留原始來源並產生衍生工作版本（P1）

作為專案維護者，我可以匯入一份 `.xlsx`，系統保留來源副本與 hash，並另存一份衍生 `.xlsx`，使後續修改不會覆蓋原始檔。

**Acceptance Criteria**：

1. 原始檔以 copy 或唯讀 snapshot 保存，匯入前後 SHA-256 相同。
2. 衍生檔具有獨立檔名與 revision，且不是原始檔的同一路徑。
3. workspace manifest 记录 `source_file`、`source_sha256`、`derived_file`、`schema_version` 與建立時間。
4. 建立衍生檔失敗時，不得刪除或修改來源檔。

### US-002：建立 workbook-v2 文字模型（P1）

作為審查者，我可以取得工作簿、工作表與儲存格的結構化 JSON，查詢每個公式、快取結果、資料型別與錯誤狀態。

**Acceptance Criteria**：

1. workbook manifest 使用 `workbook-v2` 與 `spreadsheet-fmeda` profile。
2. 每張工作表保留順序、名稱、尺寸、合併儲存格與 per-sheet JSON 位置。
3. 公式儲存格同時保留 `formula_raw`、`cached_value` 與 `calculation_status`。
4. 錯誤結果、空白與數值 0 不得互相轉換。
5. 來源檔、工作表與儲存格均可回查。

### US-003：建立公式與依賴索引（P1）

作為審查者，我可以從公式目錄查看公式原文、函數名稱、同工作表引用、跨工作表引用、外部引用與依賴範圍。

**Acceptance Criteria**：

1. `formula_catalog.csv` 每筆公式包含 formula id、sheet、cell、raw formula、cached value、status 與 function names。
2. `dependency_edges.csv` 使用 range-level edge，避免將大型範圍盲目展開成數百萬條邊。
3. 同表、跨表與外部工作簿引用必須有不同的 `reference_kind`。
4. 無法解析的公式仍保留原文並標記 risk，不得靜默丟棄。

### US-004：產生 Editor 可用的 Markdown workspace（P1）

作為審查者，我可以在現有 Markdown Block Editor 中閱讀摘要、結果、風險與審查項目，並從重要結果回查公式與來源。

**Acceptance Criteria**：

1. workspace 產生 `editor/sheets/*.md`、`editor/blocks.sidecar.json` 與 `editor/relations.json` 的相容位置。
2. Markdown 本體顯示人類可讀摘要；公式原文與依賴以可回查的 detail block 或索引連結提供。
3. 原始公式與來源快取結果預設為唯讀；review notes 可編輯。
4. 不破壞既有 Block ID、Group、TOC、sidecar 與 relations 契約。

### US-005：產生分層的審查報告（P1）

作為主管或跨部門審查者，我可以先看到工作簿狀態、主要數量、錯誤、未解析外部引用與待決策項，而不需要先閱讀全部公式。

**Acceptance Criteria**：

1. `reports/import-report.md` 說明匯入檔案、產物與統計。
2. `normalized/Step03_summary.md` 以結論、風險、證據、下一步順序呈現。
3. 任何錯誤與未解析引用可回查至 sheet + cell。
4. 摘要數字與 workbook-v2／formula catalog 使用同一份資料來源。

## 4. 非功能要求

- **Source immutability**：原始 Excel 永遠不被寫回。
- **Traceability**：所有重要結果均可回到來源檔、工作表、儲存格與公式 ID。
- **Reversibility**：任一階段失敗都能刪除 workspace 而不影響 source。
- **Scalability**：大型工作表使用 sparse extraction 與 per-sheet externalization；摘要不得強制展開全部公式。
- **Editor compatibility**：Editor 只讀取 workspace 的 Markdown 與 sidecar；FMEDA metadata 透過 extension 欄位保存。
- **Explicit uncertainty**：`cached_value`、`recalculated_value`、`error`、`unresolved` 必須分開表示。

## 5. 後續不在本 slice 內的能力

獨立公式重算器、完整資料驗證／命名範圍回存、所有 Excel metadata 的像素級保真、多人同時編輯、外部工作簿自動搜尋，以及自動決定 FMEDA 語意 mapping 均延後到後續 slices。

### US-006：受控修改與衍生 revision（P1，Slice 2）

作為審查者，我可以對明確宣告為 input 的儲存格提交 patch，加入 review note，並在來源版本與原始值檢查通過後產生新的 derived Excel revision；公式欄位仍不可修改。

**Acceptance Criteria**：

1. patch 必須使用 `fmeda-patch-v1`，並帶有 workspace 相同的 `base_source_sha256`。
2. 每一筆工作簿變更必須明確宣告 `editability=input`。
3. 公式儲存格修改必須被拒絕，且不得產生新的 revision。
4. 若 `expected_old_value` 與目前 derived workbook 不一致，必須回報 conflict 且不得產生新的 revision。
5. 成功 patch 必須寫入新的 `derived/*.rev-N.xlsx`，不得覆蓋 source snapshot 或前一個 derived revision。
6. patch manifest 必須保存來源 hash、base derived hash、new derived hash、變更儲存格與 review note 數量。
7. derived workbook 必須標記 Excel 下次開啟時重新計算；本工具不得宣稱已完成獨立公式重算。
