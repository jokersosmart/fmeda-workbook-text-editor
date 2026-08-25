# Real FMEDA Large-Workbook Integration Report

## Decision

本次 slice 的決策不是把 LibreOffice 的輸出直接視為可接受 revision，而是建立一條可以完成、可恢復、可審查的管線：原始工作簿保持唯讀，Calc 在暫存複本上重算，validator 逐張 worksheet 以 streaming 方式比較公式、結果、錯誤狀態與結構 metadata，最後以 `PASS`、`PASS_WITH_WARNINGS` 或 `FAIL` gate 住輸出。

## Input baseline

| Item | Observed result |
|---|---:|
| Source file | `RD-03-008-01FMEDAReport.xlsx` |
| Source size | 16,681,554 bytes |
| Worksheet count | 27 |
| Formula count | 597,485 |
| Merged ranges | 316 |
| Defined names | 10 |
| Uncompressed XLSX entries | 207,549,274 bytes |
| Largest worksheet XML | `sheet15.xml`, 177,287,466 bytes |
| Calculation flag | `fullCalcOnLoad=true` |

## Implementation result

LibreOffice Calc 以隔離 profile、暫存 input copy 與新的 output workbook 執行重算。原始來源沒有被寫入。重算輸出保留公式，並產生新的 output hash；之後 validator 使用 lxml tag-filtered streaming 與每張表 JSONL chunk 執行比較。

| Item | Result |
|---|---:|
| Recalculated output | `RD-03-008-01FMEDAReport.recalculated.xlsx` |
| Recalculated output size | 13,678,470 bytes |
| Worksheets validated | 27 / 27 |
| Merged-range preservation | All observed worksheet summaries reported `True` |
| Per-sheet chunks | Completed |
| Checkpoint status | `completed` |
| Formula semantic normalization | Implemented for whitespace, case, shared formulas, `FALSE()`／`TRUE()` |
| Numeric comparison | Relative／absolute tolerance `1e-12` |

## Validation result

目前真實檔案的 validation status 是 **FAIL**。這不是程式未完成，而是 validation gate 正確攔下了 Calc 重算後仍需人工審查的差異。

| Difference kind | Count | Interpretation |
|---|---:|---|
| `worksheet_metadata` | 6 | Cover、Abbreviations、BlockList、FailureRateCalcIC、FMEDA、SafetyGoalViolations 的 data validation 或 dimension metadata 發生變化 |
| `cached_value` warning | 2,774 | 多數集中於 FMEDA、BlockList 與評估表；例如 `#VALUE!` 與 `#DIV/0!` 的結果狀態轉換，不能自動當成等價 |
| `error_state` blocking | 4 | `SRAM Tran FIT` 的 4 個儲存格由 error 轉成 ok，必須人工確認 |
| `formula_raw` blocking | 0 | 公式 canonicalization 後，沒有觀察到真正的公式語意變更 |
| `unexpected_value` blocking | 0 | 極小浮點差異已由容差排除，沒有剩餘未授權的非公式值變更 |
| Total differences | 2,784 | 10 blocking、2,774 warnings |

## What this proves

本次已證明：第一，16 MB 的真實 FMEDA 可以完成 27 張 worksheet 的 streaming validation；第二，最大約 177 MB 的 worksheet XML 不必一次 materialize 成完整 cell model；第三，merged ranges、worksheet dimension、data validation、conditional formatting、tables 與 freeze panes 可以被辨識並比較；第四，Calc 重算後的 shared formula、空白、大小寫、`FALSE()`／`FALSE` 與極小浮點差異可以被分開處理；第五，validator 不會把尚未解釋的錯誤狀態直接放行。

## What this does not prove

本次尚不能宣稱 Python 已經具備獨立重算所有 FMEDA 公式的能力，也不能宣稱 LibreOffice 和 Excel 在每個函數、外部引用、extension、資料驗證與錯誤語意上完全等價。`PASS` 只能表示差異已符合目前 patch 與 validator 規則，不能取代功能安全工程師對公式與錯誤結果的判斷。

## Next decision

下一個必要的 slice 不是繼續放寬 gate，而是建立「Calc-compatible recalculation acceptance profile」：先針對 2,774 個 cached-value warnings 與 4 個 error-state changes 做分群，確認哪些是已知的原始錯誤重新分類、哪些是外部引用或 Excel extension 造成，再由工程師建立允許清單與人工審查結果。只有這些差異被解釋後，真實 FMEDA 才能從 `FAIL` 進入可接受的 `PASS_WITH_WARNINGS` 或 `PASS`。
