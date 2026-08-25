# Implementation Plan: FMEDA Workbook Text Workspace

## 1. 技術策略

本 feature 沿用既有 `mfmt/spreadsheet` rich conversion path，新增一個薄的 FMEDA core workspace service，不另起一套 Excel parser。既有 parser 負責讀取 XLSX 與輸出 per-sheet JSON；core service 負責來源保全、workbook-v2、公式／依賴／錯誤索引、可讀 Markdown 與衍生檔 manifest；Editor workspace 由可選 adapter 產生。

第一版以 read-only import 為核心，另存衍生 Excel 只做來源副本，避免在尚未完成 formula-aware patch contract 前改寫公式。第二版才開放明確允許的輸入欄位 patch，第三版才評估獨立重算。

## 2. 分層

```text
CLI / service
    ↓
FmedaCoreWorkspaceBuilder
    ├── SourceSnapshot
    ├── ExcelToJsonConverter adapter
    ├── FormulaCatalogBuilder
    ├── DependencyIndexBuilder
    ├── ReviewerMarkdownRenderer
    └── DerivedWorkbookExporter

Optional Editor adapter
    ├── EditorMarkdownRenderer
    ├── Block／Group sidecar
    ├── Relations
    └── Review workspace
```

## 3. 先後順序

1. 先補 source immutability、bundle manifest、derived copy 的契約測試。
2. 補 workbook-v2 draft 與 formula catalog 的資料契約。
3. 將現有 rich per-sheet JSON 正規化為 workbook-v2，不改動既有 generic export 的輸出契約。
4. 產生 reviewer summary 與獨立核心 Markdown；需要協作時，再由 `--adapter editor` 產生 Editor-compatible Markdown／sidecar。
5. 加入全量 validation report；先不做獨立 evaluator。
6. 完成真實 FMEDA fixture 後，才評估 input patch、derived revision 與重算。

## 4. 不可跨越的 gate

- 原始檔 hash 不一致：停止衍生輸出。
- 公式原文遺失：停止 Markdown sync 與 Excel export。
- 錯誤值被轉成 0 或空白：停止放行。
- Editor adapter sidecar 與 source revision 不一致：只允許建立 conflict report；核心 workspace 仍可獨立使用。
- 外部引用未解析：保留原始字串並標記 `unresolved`，不可自動猜測。

## 5. Phase 2 — readable core output

第二階段把核心人讀層從抽象的「Markdown summary」具體化為獨立的 `readable/` workspace：

1. `readable/index.md` 是 index-first 入口，提供來源 hash、計算模式、整體數字、審查順序、工作表索引與 machine artifact links。
2. `readable/sheets/*.md` 先呈現工作表結論與 parser status distribution，再提供輸入／常數與公式 bounded samples。
3. `readable/review-queue.md` 集中待審查項目；`readable/formula-guide.md` 解釋 formula raw、cached value、status 與 dependency 的閱讀語意。
4. `readable/manifest.json` 保存 `fmeda-readable-v1`、來源 hash、數量、路徑與展示上限；完整資料仍以 `normalized/` 為準。
5. 每張工作表最多顯示 120 筆輸入／常數與 120 筆公式，review queue 最多顯示 200 筆。上限只控制人讀視圖，不能刪除 machine artifact。
6. 優先注意工作表只能使用待審查、公式錯誤與外部依賴等 parser 訊號排序，不能宣稱為 FMEDA 失效率或安全語意判讀。

進入下一步的判準是：core-only 與 optional Editor 兩條路徑均通過 provenance／source immutability／readable output 契約測試，完整測試與格式檢查通過，且 repository 不含真實工作簿、外部檔案、ZIP 或生成的大型輸出。

下一階段可在不改變此 contract 的前提下增加 profile-specific semantic labels、可配置閱讀上限或大型 catalog 的 JSONL／搜尋索引；不得把 bounded Markdown 反向升格為計算真相。
