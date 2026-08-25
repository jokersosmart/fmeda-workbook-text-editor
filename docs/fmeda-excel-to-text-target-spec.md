# Excel 轉純文字化：重新定義後的目標規格

## 0. 本文件的決策目的

本文件重新定義「把 Excel 轉成純文字」真正要解決的問題。它不是先決定要用 Markdown、JSON 或 CSV，而是先回答：**轉換後什麼不能遺失、誰要使用、哪些內容可以修改、哪些結果需要計算、哪些證據必須能回查。**

本文件先作為目標規格與決策基線。它不代表所有後續功能已經完成，也不把尚未驗證的能力寫成既定事實。

## 1. 一體：真正要完成的一件事

### 1.1 正式目標

> **將一份 Excel 工作簿轉換成一組可獨立運作、人類可讀、程式可驗證且可回查來源的純文字工作空間；原始 Excel 永遠保留不變，所有可修改內容都透過衍生 revision 產生，不讓公式、錯誤、外部引用與計算依賴在轉換中消失。若使用 Editor，則以可選配 adapter 共用同一份模型與 provenance，提供更好的協作能力。**

因此，「純文字化」不是把 Excel 內容攤平成一個失去結構的 `.txt`，也不是只把儲存格目前看得到的值貼成 Markdown。它是將 Excel 的內容拆成不同用途、但彼此可追溯的一組文字產物。

### 1.2 完成後應該得到什麼

一次轉換至少要產生以下五種結果：

| 產物 | 主要用途 | 是否為真相來源 |
|---|---|---|
| `source manifest` | 保存原始檔名、來源 hash、版本與輸入資訊 | 是，代表來源證據 |
| `workbook.json` | 保存工作簿、工作表、儲存格、公式、結果與 metadata | 是，代表結構化計算模型 |
| `readable/` workspace | 讓審查者、主管與跨部門人員先看摘要，再逐層深入 | 否，是人類閱讀視圖 |
| `formula_catalog`／`dependency_edges` | 查詢公式、函數、引用範圍與依賴關係 | 是，代表可驗證索引 |
| `validation／review report` | 保存差異、錯誤、外部引用與人工判定 | 是，代表審查證據 |

純文字工作空間可以包含多個檔案，但不能讓它們各自形成不同的數字真相。所有摘要、驗證報告與可選的 Editor 顯示都必須回到同一份結構化模型。

## 2. 兩面：自己與別人

### 2.1 對專案維護者與 FMEDA 工程師

最重要的是不破壞原始檔、公式仍然存在、公式與結果能逐格回查、外部工作簿不會被錯誤替換，以及衍生版本能被比較。工程師需要看到公式原文、快取值、錯誤狀態、外部 link、引用範圍、工作表 metadata 與重算引擎資訊。

### 2.2 對審查者、主管與跨部門

最重要的不是一開始看到所有公式，而是先理解：這份工作簿在做什麼、主要結果是什麼、有哪些風險、哪些資料缺失、哪些結果已驗證、哪些結果仍需要工程師決定。閱讀視圖應從摘要開始，能逐層展開到工作表、區塊、儲存格、公式與來源。

### 2.3 對 Editor 使用者

Editor 是可選的協同閱讀、註記、審查與受控修改介面，不是純文字化核心的必要依賴。它不取代 Excel，也不應成為第二個未經驗證的計算引擎。若啟用 Editor，必須讓它看到與獨立核心相同的模型與 provenance，而不是複製出不同版本的內容。

## 3. 三階段：來源、工作、未來

### 3.1 過去：尊重 Excel 已有的事實

原始 Excel 可能包含公式、快取值、錯誤值、合併儲存格、命名範圍、資料驗證、條件格式、圖片、外部引用與工作表順序。轉換的第一個責任是保存這些資訊，並記錄哪些內容成功解析、哪些內容只能保留原文、哪些內容尚未能驗證。

### 3.2 現在：建立可讀、可查、可驗證的文字 workspace

現在的主工作是把 Excel 轉成一個分層 workspace：結構化 JSON 負責保存模型；Markdown 負責人類閱讀；CSV 或 JSONL 負責大型公式、依賴與差異索引；若啟用 Editor，則由 adapter 與 sidecar 負責 Block、來源位置、群組、關係與審查註記。核心流程不依賴 Editor。

### 3.3 未來：衍生 revision 與其他類似 Excel

未來不只會有這一份 FMEDA，因此不能把檔名、固定列號、固定欄位或單一工作表名稱硬寫在核心 parser。應分成共通 workbook core、FMEDA profile 與個別檔案 mapping。新的 Excel 先經過 profile detection，再判斷是 `matched`、`needs_mapping` 或 `unsupported`，不能勉強套用後製造假成功。

## 4. 四元素：Input data、Input control、Library、Output

| 元素 | 在本需求中的定義 |
|---|---|
| `Input data` | 原始 `.xlsx`、外部工作簿、Excel 中的公式、值、錯誤、metadata 與使用者提供的 mapping |
| `Input control` | 是否只讀、是否另存衍生檔、哪些欄位能編輯、重算引擎、容差、接受性 gate 與人工審查決策 |
| `Library` | workbook-v2 schema、FMEDA profile、parser、公式目錄、依賴索引、Editor sidecar、測試與既有規則 |
| `Output` | 純文字 workspace、摘要、公式／依賴索引、審查報告、patch、derived revision 與下一輪可複用的 mapping |

Output 不只是交付檔案，也要能回到下一輪成為 Library。這就是為什麼每次轉換都要留下 schema 版本、來源 hash、驗證報告與未解決事項。

## 5. 純文字化的內容層次

### 5.1 第一層：來源與結構層

這一層保存工作簿的機械事實。每個工作表要保留順序、名稱、尺寸、可見狀態、合併範圍、凍結窗格、資料驗證、條件格式、表格、命名範圍與其他可解析 metadata。每個儲存格至少要能保存座標、原始值、資料型別、公式原文與來源位置。

### 5.2 第二層：計算層

公式儲存格至少要區分以下欄位：

| 欄位 | 定義 |
|---|---|
| `formula_raw` | Excel 中原始公式文字，不能被摘要取代 |
| `formula_normalized` | 用於比較的正規化公式；不得取代原文 |
| `cached_value` | Excel 原本保存的快取結果 |
| `recalculated_value` | 由指定重算引擎產生的新結果 |
| `calculation_status` | `ok`、`error`、`unresolved_external`、`not_recalculated` 等 |
| `error_type` | 例如 `#VALUE!`、`#DIV/0!` 或其他錯誤狀態 |
| `formula_id` | 可被 catalog 與 Editor 回查的穩定識別碼 |
| `dependency_refs` | 同表、跨表、外部工作簿與範圍引用 |

`cached_value` 不得被說成最新計算結果；`#VALUE!`、空白與數值 `0` 不得被清洗成同一種狀態。

### 5.3 第三層：人類閱讀層

人類可以直接閱讀 Markdown，不需要安裝或啟用 Editor；Editor 只是可以額外提供導覽、註記與受控編輯的可選配介面。Markdown 不應該把數十萬個公式全部攤在一頁。閱讀輸出應採摘要、分區、索引與按需展開：

目前 core-only 的實體格式是 `readable/index.md`、`readable/sheets/*.md`、`readable/review-queue.md`、`readable/formula-guide.md` 與 `readable/manifest.json`。每張工作表最多展示 120 筆公式與 120 筆輸入／常數，review queue 最多展示 200 筆；這些上限只控制人讀視圖，完整公式、儲存格與 review details 仍由 `normalized/` 保存。

| 閱讀層 | 主要內容 |
|---|---|
| 主管層 | 工作簿目的、主要結果、風險、未完成事項與下一步 |
| 審查層 | 工作表摘要、關鍵結果、變更、錯誤、外部引用與待決策項 |
| 工程層 | 儲存格、公式原文、快取值、重算值、依賴、XML／metadata 證據 |
| Editor 層 | Blocks、groups、relations、review notes、source cell 與 editability |

### 5.4 第四層：證據與差異層

任何重要結論都要能從 Markdown 回到 `formula_id`、工作表、儲存格、來源檔案與來源 hash。兩個 revision 的比較要區分允許的 input 變更、公式變更、錯誤狀態變更、cached value 變更、metadata 變更與 provenance 變更。

## 6. 公式與重算的正式邊界

### 6.1 第一優先是保存公式，不是立刻取代 Excel

第一階段應保存公式原文與 Excel 快取結果，並記錄目前是否有重算。這樣即使外部工作簿、Excel extension 或重算引擎暫時不可用，原始計算真相仍然沒有消失。

### 6.2 重算分成三種層次

| 層次 | 定義 | 接受標準 |
|---|---|---|
| 保存 | 保存 Excel 原公式與快取結果 | 只要求來源忠實 |
| Excel-compatible recalc | 使用 Excel／Calc 在暫存複本重算 | 要保留引擎、版本、輸入與差異 |
| Independent evaluator | 使用系統自行重算 | 必須有函數 allowlist、錯誤語意、容差與人工 baseline 審查 |

在沒有完成第三層的函數覆蓋與 baseline 比對前，不應宣稱系統能取代 Excel。

## 7. 原始檔、衍生檔與 Editor adapter 的邊界

### 7.1 原始檔

原始 Excel 是唯讀證據來源。系統要建立 snapshot、hash 與來源 manifest，所有失敗都不能修改或刪除原始檔。

### 7.2 衍生檔

所有允許的 input 變更都應形成新的 `rev-N.xlsx`。每次變更要記錄 old value、new value、儲存格、修改者、時間、理由、影響範圍與新檔 hash。公式、來源快取值與未確認的結果預設不能在 Editor 中直接修改。

### 7.3 Editor adapter

沒有 Editor 時，核心仍應能以 CLI、Markdown、JSON、CSV／JSONL 與報告完成閱讀、驗證、patch 與衍生 revision。若啟用 Editor adapter，它只修改明確宣告為可編輯的 input 與 review notes，不應直接改寫公式，也不應把公式結果固定成數值。儲存時要產生與核心相同的 patch manifest，再由系統驗證後產生新的衍生 revision。

## 8. 外部工作簿

外部引用必須保存原始 target、link index、工作表名稱、resolved path、外部檔案 hash、mapping 狀態與 materialization 狀態。

若外部工作簿不存在、檔名對不上、hash 不明、mapping 有多個候選或 refresh 失敗，結果必須是 `unresolved` 或 `blocked`，不能用 0、空白或 synthetic data 假裝已解決。

Synthetic fixture 可以驗證程式機制，例如證明 materialization 可以讓外部 `SUMIF` 得到數值；但 synthetic 結果只能是 `review_required`，不能直接成為真實 FMEDA 結論。

## 9. 第一階段的真正範圍

第一階段不是「把所有 Excel 功能一次做完」，而是完成下列閉環：

```text
原始 Excel 唯讀
  → source manifest 與 hash
  → workbook-v2 結構化模型
  → 公式／錯誤／依賴索引
  → 人類可讀 Markdown workspace（`readable/`）
  → 可選：Editor workspace adapter
  → validation report
  → 明確標記未解析與待審查事項
```

第一階段可以先不做獨立 evaluator，也可以先不讓 Markdown 任意修改；但不能省略來源保全、公式原文、快取值、錯誤狀態、外部引用與可回查證據。

## 10. 第一階段驗收條件

| 驗收面向 | 必須達成 |
|---|---|
| 來源保全 | 原始 Excel hash 不變，且轉換失敗也不影響來源 |
| 可讀性 | 審查者、主管與跨部門可先看 `readable/index.md` 摘要，再展開工作表與完整 artifact 細節 |
| 公式保存 | 每個公式可回查原文、座標、快取值與狀態 |
| 錯誤真實性 | 錯誤、空白、0、未重算與未解析外部引用不混淆 |
| 結構保存 | 工作表順序、名稱、合併儲存格與關鍵 metadata 可回查 |
| Editor 協同（若啟用） | Markdown、sidecar、Block、relations 與 source cell 對得上；未啟用時核心仍可獨立完成流程 |
| 版本管理 | 任何修改都產生 derived revision，不覆蓋原始檔 |
| 可驗證性 | 重要結果可回到 formula_id、sheet、cell 與 source hash |
| 可複用性 | 下一份類似 Excel 能沿用共通 core，必要時只新增 profile／mapping |
| 不確定性 | 未知、衝突、待驗證與阻擋事項被明確標記，不假裝完成 |

## 11. 暫時不列為第一階段承諾

下列能力不應在第一階段被偷偷包含在「純文字化」裡：完整取代 Excel 的獨立計算器、所有 Excel metadata 的像素級複製、未經人工確認的 FMEDA 語意判斷、多人同時編輯、未明確 mapping 的外部工作簿自動搜尋，以及讓任何 Markdown 文字直接覆蓋公式。

## 12. 最終定義

最終可以用一句話定義本專案：

> **Excel 純文字化不是把儲存格內容抄成文字，而是建立一個保留來源與計算真相的可讀工作空間；JSON 保存模型，Markdown 幫人理解，CSV／JSONL 幫程式追蹤，Editor 作為可選配介面提供協作，Excel 仍是既有專業計算工具，而每一次修改都透過新的衍生 revision 與驗證報告留下可回溯證據。**

這個定義會成為後續 Spec、資料模型、API、Editor 與測試設計的上游輸入。
