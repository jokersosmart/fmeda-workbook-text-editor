# Implementation Plan: FMEDA Workbook Text Workspace

## 1. 技術策略

本 feature 沿用既有 `mfmt/spreadsheet` rich conversion path，新增一個薄的 FMEDA workspace service，不另起一套 Excel parser。既有 parser 負責讀取 XLSX 與輸出 per-sheet JSON；FMEDA service 負責來源保全、workbook-v2 draft、公式／依賴／錯誤索引、Editor workspace 與衍生檔 manifest。

第一版以 read-only import 為核心，另存衍生 Excel 只做來源副本，避免在尚未完成 formula-aware patch contract 前改寫公式。第二版才開放明確允許的輸入欄位 patch，第三版才評估獨立重算。

## 2. 分層

```text
CLI / service
    ↓
FmedaWorkspaceBuilder
    ├── SourceSnapshot
    ├── ExcelToJsonConverter adapter
    ├── FormulaCatalogBuilder
    ├── DependencyIndexBuilder
    ├── ReviewerMarkdownRenderer
    ├── EditorWorkspaceBuilder
    └── DerivedWorkbookExporter
```

## 3. 先後順序

1. 先補 source immutability、bundle manifest、derived copy 的契約測試。
2. 補 workbook-v2 draft 與 formula catalog 的資料契約。
3. 將現有 rich per-sheet JSON 正規化為 workbook-v2，不改動既有 generic export 的輸出契約。
4. 產生 reviewer summary 與 Editor-compatible Markdown／sidecar。
5. 加入全量 validation report；先不做獨立 evaluator。
6. 完成真實 FMEDA fixture 後，才評估 input patch、derived revision 與重算。

## 4. 不可跨越的 gate

- 原始檔 hash 不一致：停止衍生輸出。
- 公式原文遺失：停止 Markdown sync 與 Excel export。
- 錯誤值被轉成 0 或空白：停止放行。
- Editor sidecar 與 source revision 不一致：只允許建立 conflict report。
- 外部引用未解析：保留原始字串並標記 `unresolved`，不可自動猜測。
