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

## Slice 2 demo: controlled input patch

Slice 2 adds `FmedaPatchApplier`. A patch must declare `schema_version=fmeda-patch-v1`, match the workspace source SHA-256, declare every workbook change as `editability=input`, and may include an `expected_old_value`. Formula cells are rejected. Successful changes are written to a new derived revision and never to the source snapshot.

```powershell
python examples/run_demo.py
```

The demo creates a small three-sheet FMEDA workbook, changes `FMEDA!B1` from `0.001` to `0.002`, preserves formula `FMEDA!D1 = B1*2`, adds one reviewer note, and writes `derived/RD-03-008-01FMEDAReport.rev-002.xlsx`. It also produces `DEMO_RESULT.md`, `editor/sheets/01_FMEDA.md`, `reports/export-report.rev-002.md`, and `editor/review_notes.json` under the ignored `demo-output/` directory.

For a workspace created by `fmeda-workspace`, apply an explicit patch with:

```powershell
fmeda-patch out/fmeda-RD-03-008-01 normalized/patch.json
```

The patcher marks the derived workbook for recalculation on next Excel open with `calcMode=auto`, `fullCalcOnLoad=1`, and `forceFullCalc=1`. It does not claim to independently calculate cached formula results.

## Slice 3 demo: revision validation

Slice 3 compares a base revision and a target revision cell by cell. A declared input patch is reported as `allowed_input`; formula changes, error-state changes, unexpected values, unexpected cell presence, sheet changes, or provenance hash mismatches are blocking failures. Formula cached-value differences are warnings because this slice does not independently recalculate Excel formulas.

```powershell
fmeda-validate `
  demo-output/workspace/derived/RD-03-008-01FMEDAReport.rev-001.xlsx `
  demo-output/workspace/derived/RD-03-008-01FMEDAReport.rev-002.xlsx `
  --patch-manifest demo-output/workspace/normalized/patch_manifest.rev-002.json `
  --output demo-output/workspace/reports/validation.rev-002.md `
  --json-output demo-output/workspace/reports/validation.rev-002.json
```

A `PASS` report means all observed changes were explicitly allowed and provenance hashes matched. `PASS_WITH_WARNINGS` means only formula cached-value changes were observed. `FAIL` means a human must review the revision before it can be accepted.

## Large workbook validation

For a large workbook, use `LargeFmedaValidator` or the `fmeda-validate-large` CLI with a base and target revision. The validator reads formula and cached-value streams in read-only mode, writes one JSONL chunk per sheet, and atomically updates `checkpoint.json` after each completed sheet. If a run is interrupted, rerun with the same base and target files and completed sheets are skipped; if either file hash changes, the checkpoint is reset.

```powershell
fmeda-validate-large `
  RD-03-008-01FMEDAReport.xlsx `
  out/RD-03-008-01FMEDAReport.recalculated.xlsx `
  --output-dir out/validation-v1
```

The validator compares formula raw text, raw values, cached values, error state, cell presence, sheet order, merged ranges, data validations, conditional formatting, tables, freeze panes, and provenance hashes. Merged ranges and worksheet metadata are reported rather than silently discarded. A partial run is `INCOMPLETE`; only a complete run can be `PASS`, `PASS_WITH_WARNINGS`, or `FAIL`.

## Excel-compatible formula recalculation

`LibreOfficeRecalculator` uses a disposable headless LibreOffice profile and a temporary copy of the workbook. It never writes to the source path. The output is a new workbook with formulas preserved and formula cached results refreshed by Calc when the file is supported. The report records source hashes, output hash, formula count, cached-result count, and engine name.

```powershell
fmeda-recalculate `
  demo-output/workspace/derived/RD-03-008-01FMEDAReport.rev-002.xlsx `
  demo-output/workspace/derived/RD-03-008-01FMEDAReport.rev-002.recalculated.xlsx `
  --report demo-output/workspace/reports/recalculation.rev-002.json
```

This is deliberately an Excel-compatible recalculation path, not a claim that Python independently implements every Excel function. A future independent evaluator must be introduced only after the real FMEDA formula corpus and comparison tolerance are measured.

## Real large-workbook integration

The large-workbook path has been exercised against the real 16 MB FMEDA input. It uses XML streaming, per-worksheet JSONL chunks, resumable checkpoints, merged-range and worksheet metadata comparison, and LibreOffice Calc recalculation on a temporary copy. The original workbook is never written.

```powershell
fmeda-recalculate RD-03-008-01FMEDAReport.xlsx `
  --output out/RD-03-008-01FMEDAReport.recalculated.xlsx

fmeda-validate-large RD-03-008-01FMEDAReport.xlsx `
  out/RD-03-008-01FMEDAReport.recalculated.xlsx `
  --output-dir out/validation-v1
```

The observed baseline was 27 worksheets, 597,485 formulas, 316 merged ranges, 10 defined names, and a largest worksheet XML entry of approximately 177 MB after decompression. The completed run validated 27/27 worksheets and reported all observed merged ranges as preserved. A `FAIL` status means the validator found blocking differences that still require engineering review; it is not silently converted to PASS.

See `docs/real-fmeda-integration-report.md` for the observed result and remaining acceptance work. The real workbook and generated validation chunks are intentionally excluded from version control.

## External workbook loading and materialization

Some FMEDA formulas use references such as `[1]BlockList!`. The recalculator can now accept supplied external workbooks by exact basename mapping and record the resolved file hash. `bind` keeps the external relationship pointing at the supplied file for Excel-compatible workflows. `materialize` copies the referenced external worksheet into a temporary internal sheet such as `EXT_BlockList` and rewrites only the temporary calculation copy, which is the recommended mode when Calc cannot evaluate the original `[1]` link-index syntax.

```powershell
fmeda-recalculate `
  RD-03-008-01FMEDAReport.xlsx `
  out/RD-03-008-01FMEDAReport.with-external.xlsx `
  --external-workbook C:/path/SM2734_HWS_SA_FMEDA_0.2chk.xlsx `
  --external-mode materialize `
  --report out/external-recalculation.json
```

The source FMEDA and supplied external workbook are never written. If the exact external basename cannot be resolved, the command fails closed instead of guessing. The report records `resolved_path`, external SHA-256, source SHA-256, materialized sheet name, and the fact that formula rewrites were limited to the temporary copy.

A reproducible synthetic demonstration is available at:

```powershell
python examples/run_external_link_demo.py
```

It starts four cells with `#VALUE!`, loads a clearly synthetic `SM2734_HWS_SA_FMEDA_0.2chk.xlsx`, materializes `BlockList` as `EXT_BlockList`, and recalculates `T2`, `W2`, `T3`, and `W3` to numeric values. The demonstration validates the mechanism only; its values are not the real FMEDA result. The real external workbook must still be supplied and reviewed before the four production cells can be accepted.


## External recalculation acceptance profile

A numeric result after external materialization is not automatically accepted. `FmedaAcceptanceProfile` classifies each target cell as `accepted`, `review_required`, or `blocked`. Synthetic fixtures remain `review_required`; unresolved links, missing hashes, remaining errors, formula changes, and incomplete provenance are `blocked`. A production result becomes `accepted` only when the external workbook hash is recorded and a reviewer decision manifest contains both an identified reviewer and a rationale.

```powershell
fmeda-acceptance `
  RD-03-008-01FMEDAReport.xlsx `
  out/RD-03-008-01FMEDAReport.with-external.xlsx `
  --recalc-report out/external-recalculation.json `
  --output out/external-acceptance.md `
  --json-output out/external-acceptance.json
```

The same gate can be attached to the ordinary revision validator with `--recalc-report`. The Markdown report contains a manager summary, a reviewer table, engineer-level formula evidence, and provenance. All views use the same underlying decision; they only change the amount of explanation.


## Complete thinking-rule catalog

The repository now includes the complete source-derived thinking-rule catalog from the supplied decision database. All 45 numbered reports were read in full. The catalog contains 235 habit-definition lines and 185 program-rule candidate lines, with zero extraction errors. The source-derived JSON is under `resources/thinking/`, the reading checklist and rule-engine contract are under `docs/thinking/`, and the compiled catalog preserves every rule's report number and source filename.

Compile the catalog and evaluate a project decision context with:

```powershell
thinking-rules compile `
  --catalog resources/thinking/thinking_rule_catalog.json `
  --output resources/thinking/compiled_program_rules.json

thinking-rules evaluate `
  --catalog resources/thinking/thinking_rule_catalog.json `
  --context demo-output/thinking-rules/fmeda-context.json `
  --json-output demo-output/thinking-rules/evaluation.json `
  --markdown-output demo-output/thinking-rules/evaluation.md
```

The engine deliberately has three execution levels: `deterministic_check`, `human_review`, and `assistive_prompt`. Subjective, ambiguous, or value-laden habits remain review-required or prompt-only. The engine is a decision aid and does not replace domain review or the user's final judgment.


The one-to-one habit catalog is `resources/thinking/habit_program_catalog.json`. It contains one program-rule candidate for each of the 235 habit-definition lines across all 45 reports. Compile it separately when the goal is complete habit coverage:

```powershell
thinking-rules compile `
  --catalog resources/thinking/habit_program_catalog.json `
  --output resources/thinking/compiled_habit_program_rules.json
```

The original 185-candidate catalog is retained because report-level candidates and habit-level candidates answer different questions. The former preserves the report's explicit candidate rules; the latter guarantees that no habit definition is silently omitted. Every habit-level rule remains `candidate` and `disabled_until_owner_confirmation` until the owner confirms its operational meaning.
