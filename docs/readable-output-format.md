# Readable Output Format

## 1. 文件目的

`readable/` 是 FMEDA core workspace 的**人類閱讀層**。它讓審查者、主管與跨部門人員先看到工作簿狀態、風險訊號與下一步，再按工作表深入到公式與來源證據；它不建立另一套數字真相，也不需要 Markdown Editor 才能使用。

> `normalized/` 保存機器可驗證的完整資料；`readable/` 提供有界、可導覽的閱讀視圖；`editor/` 若啟用，只是額外的協作 adapter。

## 2. 產物契約

| 路徑 | 作用 | 真相來源或閱讀視圖 |
|---|---|---|
| `readable/index.md` | 工作簿閱讀入口、整體數字、審查順序、工作表索引與 artifact links | 閱讀視圖 |
| `readable/sheets/<index>_<sheet>.md` | 每張工作表的結論、狀態分布、輸入樣本、公式樣本與追溯路徑 | 閱讀視圖 |
| `readable/review-queue.md` | 依類型與狀態整理待審查項目，最多展示前 200 筆 | 閱讀視圖；完整資料仍在 JSON |
| `readable/formula-guide.md` | 解釋 `formula_raw`、`cached_value`、`calculation_status` 與依賴欄位 | 閱讀說明 |
| `readable/manifest.json` | readable schema、來源 hash、數量、限制與各頁路徑 | readable metadata |
| `normalized/Step03_workbook.json` | 完整 workbook-v2 結構化模型 | 機器真相 |
| `normalized/formula_catalog.csv` | 所有公式的原文、快取值、狀態、函數與依賴 | 機器真相 |
| `normalized/dependency_edges.csv` | range-level 同表、跨表與外部依賴 | 機器真相 |
| `normalized/review_items.json` | 完整 review item 與 details | 機器真相 |

## 3. 建議閱讀順序

第一步先看 `readable/index.md`，確認來源檔、`source_sha256`、計算模式、公式數、依賴數與待審查項目。第二步看 `review-queue.md` 與 `formula-guide.md`，先處理錯誤、外部引用、未計算結果與其他未決事項。第三步按工作表頁閱讀結論與樣本。第四步才從 `formula_id`、來源儲存格與相對連結回查完整 JSON、CSV 與 review item。

工作表索引中的「優先注意工作表」只依核心 parser 產生的待審查、公式錯誤與外部依賴訊號排序。它**不是** FMEDA 失效率、安全關鍵性或工程語意的自動判讀；真正的工程結論仍由 FMEDA 工程師與指定審查者負責。

## 4. 工作表頁面

每張工作表頁面依序包含四個區段：

1. **先看結論**：工作表順序、尺寸、合併儲存格數、公式數、輸入／常數數、待審查數、外部引用數、公式錯誤訊號數與已辨識的函數名稱。
2. **關鍵輸入／常數**：以來源儲存格、值、資料型別與 `source_sha256` 顯示前 120 筆非公式內容。
3. **公式計算摘要**：以來源儲存格、`formula_raw`、`cached_value`、狀態與 `formula_id` 顯示前 120 筆公式。
4. **如何追溯**：連回完整工作表 JSON、公式 catalog、依賴索引、review queue 與來源 hash。

Markdown 表格中的 pipe 與換行會被轉義或壓成單行，以免來源內容破壞表格結構。原始內容仍保留在 normalized JSON／CSV。

## 5. Bounded detail policy

大型工作表不會把數十萬筆公式全部攤在 Markdown。每張工作表預設最多展示 120 筆輸入／常數與 120 筆公式，review queue 預設最多展示 200 筆。頁面會明確標示實際顯示數量與完整資料路徑；這些限制只適用於閱讀視圖，不會刪除或截斷 normalized artifact。

| 類別 | readable 預設上限 | 完整資料 |
|---|---:|---|
| 每張工作表公式樣本 | 120 | `normalized/formula_catalog.csv` 與 per-sheet JSON |
| 每張工作表輸入／常數樣本 | 120 | per-sheet JSON |
| review queue 展示項目 | 200 | `normalized/review_items.json` |

## 6. 計算語意與接受邊界

`cached_value` 代表來源 Excel 保存的快取結果，不代表 readable builder 或 Python 已重新計算。`error`、`not_calculated`、空白與數值 `0` 必須分開解讀。外部工作簿引用如果仍是 `unresolved`，必須保留原始公式與 review item，不得用 0、空白或 synthetic data 假裝已解決。

若日後使用 Calc 或其他 Excel-compatible engine 重算，重算結果必須寫入新的衍生 revision，並以既有 validation／acceptance contract 判斷。readable 頁面可以引用重算報告，但不會自行把快取值改成「已驗證結果」。

## 7. Provenance contract

所有 readable 頁面與 `readable/manifest.json` 都沿用 core workspace 的 `source_sha256`。工作表公式以 `formula_id`、`source_cell` 與 `formula_raw` 回查；完整結果以 normalized artifact 為準。Editor adapter 若啟用，必須從同一份 core model 與 provenance 產生，不能建立第二套公式或數字來源。

## 8. 未來泛化

readable renderer 不應依賴目前 FMEDA 的固定列號、欄位名稱或單一工作表。共通層只處理 workbook、sheet、cell、formula、dependency、review 與 provenance；未來的 FMEDA profile 或其他 Excel mapping 可以提供語意標籤，再由 readable layer 顯示。若 profile 未匹配，輸出必須保留 `needs_mapping` 或 `unsupported` 等狀態，而不是產生看似完整的假結論。
