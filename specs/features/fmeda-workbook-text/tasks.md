# Tasks: FMEDA Workbook Text Workspace

## Slice 0 — 基線與契約

- [ ] 0-1. 修正 `ExcelToJsonConverter` 測試與實作對 `extract_images`／`extract_images_enabled` 的設定契約。
- [x] 0-2. 建立 source immutability contract test。
- [x] 0-3. 建立 derived copy 與 source hash manifest contract test。

## Slice 1 — Read-only FMEDA workspace

- [x] 1-1. 建立 `FmedaWorkspaceBuilder` 與 workspace manifest。
- [x] 1-2. 將既有 rich per-sheet JSON 正規化為 `workbook-v2` draft。
- [x] 1-3. 建立 formula catalog，保留 raw formula、cached value、status 與 function names。
- [x] 1-4. 建立 range-level dependency edges，分辨 same-sheet、cross-sheet、external。
- [x] 1-5. 建立 review items，保留 formula error、unresolved external reference、unsupported feature。
- [x] 1-6. 產生 reviewer／management summary Markdown。
- [x] 1-7. 產生 Editor-compatible Markdown 與 sidecar metadata。

## Slice 2 — Editor 協同與受控修改

- [x] 2-1. 在 Block sidecar 中建立 `source_cell`、`formula_id`、`source_revision` 與 `editability`。
- [x] 2-2. 允許只修改明確宣告 `editability=input` 的儲存格。
- [x] 2-3. 建立 patch manifest、expected-value check 與 source revision conflict detection。
- [x] 2-4. 將 review notes 寫入 Editor review sidecar。
- [x] 2-5. 編輯成功後建立下一版 derived XLSX，不覆蓋原始來源。

## Slice 3 — 全量驗證與回存

- [x] 3-1. 建立全量 sheet／cell／formula／value／error／provenance diff。
- [x] 3-2. 建立 `fmeda-validate` CLI。
- [ ] 3-3. 擴充 merged cells、defined names、number formats、data validations 的驗證。
- [x] 3-4. 建立 derived export report 與可回復 revision。

## Slice 4 — Profile 泛化

- [ ] 4-1. 建立共通 workbook profile contract。
- [ ] 4-2. 建立 FMEDA profile detection 與 mapping 檔格式。
- [ ] 4-3. 對第二份類似 Excel 執行 `matched`／`needs_mapping`／`unsupported` 路由。
- [ ] 4-4. 建立跨檔案 regression corpus。

## Slice 5 — 獨立重算（延後）

- [ ] 5-1. 盤點真實 FMEDA 使用的函數與運算型態。
- [ ] 5-2. 建立 evaluator contract 與誤差／錯誤語意規則。
- [ ] 5-3. 以 Excel 結果作為 baseline，執行全量結果比對。
- [ ] 5-4. 只有在差異可解釋且通過人工審查後，才開放獨立重算。

## Slice 6 — 真實大型 FMEDA 與公式重算

- [x] 6-1. 建立真實 FMEDA inventory 基線：27 sheets、597,485 formulas、316 merged ranges、10 defined names。
- [x] 6-2. 建立 per-sheet JSONL chunks 與 atomic checkpoint，支援中斷後 resume。
- [x] 6-3. 比較公式原文、raw value、cached value、error state 與 cell presence。
- [x] 6-4. 比較 merged ranges、worksheet dimension、data validations、conditional formatting、tables 與 freeze panes。
- [x] 6-5. 建立 LibreOffice Calc 隔離 profile 重算與新檔輸出。
- [ ] 6-6. 以真實 FMEDA 的全部公式函數建立獨立 evaluator allowlist 與結果 tolerance。
- [x] 6-7a. 完成真實 16 MB 工作簿的 Calc 重算、27/27 worksheet validation 與可恢復 checkpoint。
- [ ] 6-7b. 完成 cached-value warnings、error-state changes 與 metadata changes 的工程師人工抽樣審查。
