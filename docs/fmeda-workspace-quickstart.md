# FMEDA Workspace Quickstart

## 目的

`fmeda-workspace` 會把一份 `.xlsx` 轉成 source-safe FMEDA core workspace。原始檔不會被覆蓋；系統會複製來源到 `source/`，建立同內容的衍生工作簿到 `derived/`，並產生 workbook-v2、公式目錄、依賴索引、審查清單與人類可讀 Markdown。Editor workspace 是可選 adapter，不是核心必要依賴。

## 執行

```powershell
fmeda-workspace `
  RD-03-008-01FMEDAReport.xlsx `
  --output-dir out/fmeda-RD-03-008-01
```

在 Windows PowerShell 中，反引號是換行符號；也可以寫成單行：

```powershell
fmeda-workspace RD-03-008-01FMEDAReport.xlsx -o out/fmeda-RD-03-008-01
```

## 產物

```text
out/fmeda-RD-03-008-01/
├── manifest.json
├── source/
│   └── RD-03-008-01FMEDAReport.xlsx       # 唯讀來源快照
├── derived/
│   └── RD-03-008-01FMEDAReport.rev-001.xlsx # 可供後續工作的新檔
├── normalized/
│   ├── Step03_workbook.json
│   ├── Step03_summary.md
│   ├── formula_catalog.csv
│   ├── dependency_edges.csv
│   ├── review_items.json
│   └── sheets/*.json
└── reports/
    └── import-report.md

# 只有使用 --adapter editor 時才會額外產生 editor/ 目錄：
# index.md、sheets/*.md、blocks.sidecar.json、relations.json
```

## 與 Editor 搭配（可選）

核心不需要 Editor 即可完成純文字化、閱讀與驗證。若要使用 Markdown Block Editor，再以 adapter 模式產生 `editor/`：

```powershell
fmeda-workspace RD-03-008-01FMEDAReport.xlsx `
  --output-dir out/fmeda-RD-03-008-01-editor `
  --adapter editor
```

此時 `editor/index.md` 是入口，`editor/sheets/*.md` 是依工作表拆分的文件。每個公式明細以 `formula_id`、來源儲存格與原始公式標記；完整公式仍以 `normalized/formula_catalog.csv` 為準。

`blocks.sidecar.json` 目前把來源摘要與公式明細標成唯讀，並保留 `source_cell`、`formula_id`、`source_revision` 與 `editability`。這個切片先不開放普通 Markdown 編輯覆蓋公式；後續的 review notes 與 input patch 必須另外經過權限與來源版本檢查。

## 目前的計算語意

目前 workspace 使用來源 Excel 保存的 cached values，`calculation_mode` 為 `source_cached_values`。這不等同於本工具重新計算。公式、空白、數值 0、`#DIV/0!`、`#VALUE!` 與未解析外部引用會分別保存並列入審查項目。

## 進入下一個 slice 的條件

當 source hash、工作表順序、儲存格集合、公式原文、快取結果與錯誤狀態都有全量驗證後，才開放 input patch 與新的 derived revision。獨立 evaluator 仍然延後，必須以實際 FMEDA 結果做可解釋的差異比對後才能啟用。
