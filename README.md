# FMEDA Workbook Text Editor

這個 repository 是 `project2md_controll` 中 FMEDA Workbook Text Workspace vertical slice 的獨立版本。它的目標不是把 Excel 簡單轉成一個失去計算語意的 `.txt`，而是建立一個**原始檔保全、衍生檔工作、Markdown Editor 協同審查**的工作區。

> 原始 Excel 是來源證據；衍生 Excel 是工作版本；Markdown Editor 是理解、審查、註記與導覽入口。

## Current slice

目前完成的是 Slice 1：read-only FMEDA workspace。它會保留來源 Excel 的 SHA-256，另存一份 derived workbook，並輸出 `workbook-v2` JSON、公式目錄、依賴索引、審查清單、分層摘要與 Editor-compatible Markdown／sidecar。

第一階段不會把公式改寫成 Python，也不會讓普通 Markdown 編輯覆蓋公式。Excel 快取值會標記為 `source_cached_values`；`error`、空白、0、未計算與外部引用未解析狀態會分開保存。

## Development tree

```text
FMEDA Workbook Text Workspace
├── Slice 0：Source 保全與基線
├── Slice 1：Read-only FMEDA Workspace       ← current
│   ├── workbook-v2 JSON
│   ├── formula_catalog.csv
│   ├── dependency_edges.csv
│   ├── review_items.json
│   ├── Reviewer／主管摘要 Markdown
│   └── Editor Markdown + sidecar
├── Slice 2：Editor 協同與受控修改
├── Slice 3：全量驗證與衍生回存
├── Slice 4：Profile 泛化到其他 Excel
└── Slice 5：獨立重算：最後才做
```

完整圖檔請查看 [`docs/fmeda-development-tree.png`](docs/fmeda-development-tree.png)，原始 Mermaid 位於 [`docs/fmeda-development-tree.mmd`](docs/fmeda-development-tree.mmd)。

## Install

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
```

## Run

```powershell
fmeda-workspace RD-03-008-01FMEDAReport.xlsx --output-dir out/fmeda-RD-03-008-01
```

The source workbook is never written to. The output contains:

```text
out/fmeda-RD-03-008-01/
├── manifest.json
├── source/                         # source snapshot
├── derived/                        # derived workbook revision
├── normalized/
│   ├── Step03_workbook.json
│   ├── Step03_summary.md
│   ├── formula_catalog.csv
│   ├── dependency_edges.csv
│   ├── review_items.json
│   └── sheets/*.json
├── editor/
│   ├── index.md
│   ├── sheets/*.md
│   ├── blocks.sidecar.json
│   └── relations.json
└── reports/import-report.md
```

## Test

```powershell
python -m pytest -q
```

The contract tests cover source immutability, derived copy creation, formula preservation, same-sheet／cross-sheet references, unresolved external references, formula errors and Editor provenance metadata.

## Scope and safety

This repository deliberately does not include the user's original FMEDA workbook. The source file should be supplied locally when running the command. The focused repository also does not include unrelated uncommitted files from the original project.

The next implementation slice should add explicit input patch permissions, review notes, source revision conflict detection and derived `rev-N` export. Independent formula recalculation remains deferred until a real FMEDA corpus establishes the required function set and acceptable comparison rules.
