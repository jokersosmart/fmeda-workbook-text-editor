# External Workbook Materialization Report

## Purpose

本次實驗的目的，是確認程式能否在不修改原始 FMEDA 的前提下，載入一個外部 `BlockList` 工作簿，將它 materialize 成暫存內部 worksheet，交給 LibreOffice Calc 重算原本使用 `[1]BlockList!` 的公式。

## Important boundary

本報告使用的是明確標記的 synthetic external workbook，檔名模擬為 `SM2734_HWS_SA_FMEDA_0.2chk.xlsx`。它只用來驗證載入、mapping、materialization 與重算路徑；其中的 `AO`、`C2`、`C3` 數值不是實際 FMEDA 資料，不能解讀成產品安全分析結論。

## Real workbook experiment

Input 是使用者提供的真實 `RD-03-008-01FMEDAReport.xlsx`。系統先讀取原始 external relationship，再以 exact basename 對應 synthetic workbook。重算流程只在暫存複本新增 `EXT_BlockList`，並把暫存複本中的 `[1]BlockList!` 改寫為 `EXT_BlockList!`；原始檔與 synthetic external workbook 都不會被寫回。

| Cell | Original cached result | Materialized synthetic result | Formula role |
|---|---:|---:|---|
| `SRAM Tran FIT!T2` | `#VALUE!` | `4096` | External SUMIF root |
| `SRAM Tran FIT!W2` | `#VALUE!` | `0.96484375` | Downstream calculation from T2 |
| `SRAM Tran FIT!T3` | `#VALUE!` | `8192` | External SUMIF root |
| `SRAM Tran FIT!W3` | `#VALUE!` | `3.7421875` | Downstream calculation from T3 |
| `SRAM Tran FIT!X2` | `#VALUE!` | `0.2412109375` | External denominator branch |
| `SRAM Tran FIT!X3` | `#VALUE!` | `1.87109375` | External denominator branch |

The materialized calculation copy contains these formula forms:

```excel
=SUMIF(EXT_BlockList!R10:R49,"Si MOS: High speed SRAM, FIFO",EXT_BlockList!AO10:AO49)
=SUMIF(EXT_BlockList!R10:R49,"Si MOS: Digital circuits, Micros, DSP",EXT_BlockList!AO10:AO49)
=(U2+V2)*T2/1024/1024
=(U3+V3)*T3/1024/1024
=W2/EXT_BlockList!C3
=W3/EXT_BlockList!C2
```

## What this proves

The external-link implementation can discover a standard Excel `externalReference`, resolve it by an exact basename, record the external workbook hash, add a temporary internal worksheet, rewrite only the calculation copy, and allow Calc to compute the four previously failing production cells as numeric values. It also resolves the two related `X2`／`X3` denominator branches in this synthetic experiment.

The real 16 MB workbook remains compatible with the streaming recalculation counter: the report recorded 597,485 formulas and 597,485 cached results after the materialized run, with the original source hash unchanged before and after execution.

## What this does not prove

This experiment does not prove that `4096`, `8192`, `0.96484375`, `3.7421875`, `0.2412109375` or `1.87109375` are the correct production results. Those values came from the synthetic external workbook. It also does not prove that direct LibreOffice evaluation of the original `[1]` syntax will work; the direct bind test produced `#NAME?`, which is why `materialize` is the recommended Calc path.

To accept the production result, the actual `SM2734_HWS_SA_FMEDA_0.2chk.xlsx` must be supplied. The system should then compare its hash, `BlockList!R10:R49`, `BlockList!AO10:AO49`, `BlockList!C2`, and `BlockList!C3`, rerun the materialized calculation, and have an FMEDA engineer review the resulting T2／T3／W2／W3／X2／X3 values.

## Decision

The four original `#VALUE!` states should remain blocked until the actual external workbook is loaded. The implementation is ready to perform that load; the acceptance decision is not yet ready because the real external data is still unavailable.
