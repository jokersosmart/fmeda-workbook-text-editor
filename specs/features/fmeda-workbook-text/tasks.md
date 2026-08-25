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


## Slice 7 — 外部工作簿載入與重算解錯

- [x] 7-1. 讀取標準 `externalReference`／`externalLink` relationship，依 exact basename 建立 resolved／unresolved mapping。
- [x] 7-2. 在暫存複本中保留 direct bind 模式，並建立 `EXT_<Sheet>` internal-sheet materialization 模式。
- [x] 7-3. 將 external workbook path、SHA-256、materialized sheet、公式改寫範圍寫入 recalculation report。
- [x] 7-4. 以 synthetic external BlockList 對真實 FMEDA 執行 Calc materialization，確認 T2、W2、T3、W3 與 X2、X3 可產生數值。
- [ ] 7-5. 取得真正的 `SM2734_HWS_SA_FMEDA_0.2chk.xlsx`，以實際 BlockList 值重跑並完成 FMEDA 工程師人工審查。


## Slice 8 — External recalculation acceptance profile

- [x] 8-1. 定義 `accepted`、`review_required`、`blocked` 三段式 gate 與 evidence contract。
- [x] 8-2. 實作 `FmedaAcceptanceProfile`，接續 recalc report、external hash、formula equivalence 與 reviewer decision manifest。
- [x] 8-3. 加入 `fmeda-acceptance` CLI，並可由 `fmeda-validate --recalc-report` 觸發同一份 acceptance gate。
- [x] 8-4. 產生 manager、reviewer、engineer 三層視圖；視圖不得改變判定結果。
- [x] 8-5. 以真實 16 MB FMEDA + synthetic external fixture 驗證：6 個目標結果為 `review_required`，不自動 accepted。
- [ ] 8-6. 取得真正 external workbook 後，以 production source kind 重跑並完成 FMEDA 工程師 decision manifest。


## Slice 9 — Complete thinking-rule catalog and preflight

- [x] 9-1. 逐份讀完 ZIP 中 45 份編號決策報告，建立來源與報告編號對照。
- [x] 9-2. 建立完整閱讀 checklist，保存核心問題、習慣定義、判斷順序、觸發、證據、不確定性、停損與輸出欄位。
- [x] 9-3. 正規化 235 條習慣定義行與 185 條程式規則候選行，保留零讀取錯誤與跨報告歧義。
- [x] 9-4. 建立 `joker-thinking-program-rules-v1` compiled catalog，逐條保留 source_file、source_text 與 execution level。
- [x] 9-5. 實作 `deterministic_check`、`human_review`、`assistive_prompt` 三種執行邊界；主觀或歧義內容不得自動升級。
- [x] 9-6. 實作 `thinking-rules compile` 與 `thinking-rules evaluate` CLI，輸出同一 evaluation object 的 JSON 與 Markdown。
- [x] 9-7. 建立 45 報告、185 規則、目標缺失、高代價不可逆與完整低風險 context 的契約測試。
- [ ] 9-8. 由使用者逐項確認各報告中的歧義、量化門檻與跨報告衝突，才可將需要人判斷的候選升級為更嚴格的 deterministic rule。
