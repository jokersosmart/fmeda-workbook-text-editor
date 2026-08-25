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

### US-007：驗證 derived revision 的可追溯差異（P1，Slice 3）

作為審查者，我可以比較兩個 derived workbook revision，知道 input 變更是否符合已提交的 patch，並在公式、錯誤狀態、非預期值或來源 hash 異常時阻止接受該 revision。

**Acceptance Criteria**：

1. validator 必須逐格比較公式原文、raw value、cached value、error state 與 cell presence。
2. 只有 patch manifest 中明確列出的 input 變更可以標記為 `allowed_input`。
3. 公式原文、錯誤狀態、非預期值、工作表存在／順序或 provenance hash 差異必須標記為 blocking。
4. 公式 cached value 差異可以標記為 warning，因為本 slice 不宣稱完成獨立重算。
5. `fmeda-validate` 必須輸出 Markdown report，並以 `PASS`、`PASS_WITH_WARNINGS` 或 `FAIL` 結束。
6. `FAIL` 時 CLI 必須以非零狀態結束，避免未經審查的 revision 被自動接受。

### US-008：大型 FMEDA 可恢復驗證（P1，大檔 slice）

作為審查者，我可以對大型 FMEDA 的兩個 revision 執行分頁、分批且可恢復的 validation；即使中途停止，也不需要從第一張工作表重新開始，並且能知道哪些工作表已完成。

**Acceptance Criteria**：

1. validator 必須以 read-only workbook streams 處理工作表，避免把整個大型 workbook 展開成單一 cell model。
2. 每張工作表完成後，必須以 atomic write 更新 checkpoint 與該工作表 JSONL chunk。
3. base／target hash 改變時，舊 checkpoint 不得被誤用，必須重新開始。
4. 未完成的執行狀態必須是 `INCOMPLETE`，不得被誤判為 `PASS`。
5. merged ranges、data validations、conditional formatting、tables、freeze panes 與 worksheet dimension 必須被保留並可比較。

### US-009：Excel-compatible 公式重算（P1，Calc slice）

作為 FMEDA 工程師，我可以在不修改原始來源的前提下，使用 Excel-compatible engine 產生一份重算後的工作簿，並知道公式快取結果是否已由該 engine 寫入。

**Acceptance Criteria**：

1. 重算必須在暫存複本與隔離 profile 上執行，不得直接寫入 source workbook。
2. 輸出工作簿必須保留公式原文，並另外產生 source／output hash 與 formula／cached-result 統計。
3. 重算引擎與執行狀態必須記錄；不可以把 LibreOffice 結果宣稱為 Python 獨立 evaluator 結果。
4. 外部引用、Excel 不支援的 extension 或重算錯誤必須保留為風險狀態，交由人工審查。


### US-010：真實大型工作簿 integration（P1，Calc + streaming validation slice）

作為審查者，我可以對實際的大型 FMEDA 工作簿執行 Calc-compatible 重算與 27 張工作表的可恢復 validation，並看到哪些差異已經可解釋、哪些仍需要工程師確認。

**Acceptance Criteria**：

1. 16 MB 等級工作簿可以完成所有 worksheet 的 streaming validation，不因單一 worksheet XML 約 177 MB 而要求整本 workbook materialization。
2. formula cache 統計使用 worksheet XML streaming，不以四個 openpyxl workbook 物件同時掃描大型工作簿。
3. shared formula follower 以 anchor 與相對位置展開後比較；空白、大小寫、`FALSE()`／`FALSE` 等等價語法不得被誤判為 formula change。
4. 數值差異使用明確的相對／絕對容差；容差不適用於錯誤字串、公式原文、cell presence 或 metadata。
5. 27/27 worksheet 完成、checkpoint 狀態為 `completed`，且 merged ranges 均保留為可比較 metadata。
6. Calc 重算後的 cached-value 變化與 error-state 變化必須保留；未解釋差異應維持 `FAIL` 或 warning，不得自動轉成 `PASS`。
7. 真實工作簿、重算檔與大型 JSONL chunks 不得進入版本控制；repository 只保存程式、契約、文件與去識別化的執行摘要。

**Observed integration result**：27/27 worksheets completed；公式語意 blocking changes 為 0；浮點／非公式誤報在容差後為 0；剩餘 10 個 metadata／error-state blocking changes 與 2,774 個 cached-value warnings，已記錄於 `docs/real-fmeda-integration-report.md`。


### US-011：外部工作簿 mapping 與可驗證重算（P1，external-link slice）

作為 FMEDA 工程師，我可以提供公式所引用的外部工作簿，系統在暫存複本中解析並 materialize 外部 link mapping，再交給 Calc 重算；若外部工作簿不存在或 mapping 不明確，系統必須保持阻擋狀態，不得用模擬值冒充真實結果。

**Acceptance Criteria**：

1. 系統可從 XLSX relationship 讀出原始 external target、external link index 與 sheet name。
2. 使用者提供的 external workbook 必須依檔名或明確 link index 對應；找不到或多重候選時不得猜測。
3. 綁定只發生在暫存複本；原始 FMEDA 與使用者提供的外部工作簿都不可被寫回。
4. 重算報告必須保存原始 target、實際 resolved path、external workbook SHA-256、mapping status 與 materialization status。
5. synthetic fixture 只能驗證流程能否解開外部公式，不得被標記為真實 FMEDA 結論。
6. 真實外部工作簿尚未提供時，`unresolved`／`refresh_error` 必須維持 `blocked`，不能把 `#VALUE!` 自動變成可信的 0。
7. 若外部工作簿成功載入，`T2`、`T3`、`W2`、`W3` 的 cached result 仍需經 validator 與工程師抽樣確認後，才可解除 blocking。
