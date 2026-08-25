# SRAM Tran FIT 四個 Error-State Changes 根因分析

## 結論先行

這 4 個 error-state changes 並不是 4 個彼此獨立的問題，而是 **2 個外部引用根節點 + 2 個下游公式節點**：

```text
外部工作簿 [1]BlockList
        ├── SRAM Tran FIT!T2  →  #VALUE! → 0
        │                         └── SRAM Tran FIT!W2  →  #VALUE! → 0
        └── SRAM Tran FIT!T3  →  #VALUE! → 0
                                  └── SRAM Tran FIT!W3  →  #VALUE! → 0
```

目前最可信的根因是：**原始工作簿的 `SUMIF` 公式依賴一個未能成功更新的外部工作簿；原始檔保留了 `#VALUE!` 快取錯誤，而 LibreOffice Calc 在沒有取得外部資料時，將相同公式的 `SUMIF` 結果寫成 0，導致下游乘法公式由錯誤傳播改成 0。** 這不是已證明的「Excel 正確結果」，而是一次 Calc 重算在缺少外部來源時的行為。

## 四格其實是兩條依賴鏈

| 儲存格 | 原始公式 | 原始 cached value | Calc cached value | 角色 |
|---|---|---:|---:|---|
| `T2` | `SUMIF([1]BlockList!R10:R49,"Si MOS: High speed SRAM, FIFO",[1]BlockList!AO10:AO49)` | `#VALUE!` | `0` | 外部引用根節點 |
| `W2` | `(U2+V2)*T2/1024/1024` | `#VALUE!` | `0` | T2 的下游計算 |
| `T3` | `SUMIF([1]BlockList!R10:R49,"Si MOS: Digital circuits, Micros, DSP",[1]BlockList!AO10:AO49)` | `#VALUE!` | `0` | 外部引用根節點 |
| `W3` | `(U3+V3)*T3/1024/1024` | `#VALUE!` | `0` | T3 的下游計算 |

原始與重算檔的 4 個公式文字相同；變化發生在 cached result 與 cell error type，而不是公式本體。

## 證據一：公式直接指向外部工作簿

兩個根節點 `T2` 與 `T3` 都使用 `[1]BlockList`，而不是本工作簿內的 `BlockList`：

```excel
=SUMIF([1]BlockList!R10:R49,"Si MOS: High speed SRAM, FIFO",[1]BlockList!AO10:AO49)
=SUMIF([1]BlockList!R10:R49,"Si MOS: Digital circuits, Micros, DSP",[1]BlockList!AO10:AO49)
```

外部連結的 XML relationship 指向：

```text
<original workstation path>/SM2734_HWS_SA_FMEDA_0.2chk.xlsx
```

本次可用附件與 sandbox 中沒有找到這個檔案，因此 Calc 重算時沒有可供載入的外部 `BlockList`。

## 證據二：原始檔明確標記外部連結 refresh error

原始檔 `xl/externalLinks/externalLink1.xml` 具有：

```xml
<sheetData sheetId="0" refreshError="1" />
```

這表示原始工作簿本身已保存一個「外部資料更新失敗」訊號。原始檔中 `T2`、`T3` 的 XML 儲存格也是：

```xml
<c r="T2" ... t="e"><f>SUMIF(...)</f><v>#VALUE!</v></c>
<c r="T3" ... t="e"><f>SUMIF(...)</f><v>#VALUE!</v></c>
```

因此原始 `#VALUE!` 不是我們在轉換過程中創造的，而是原始 Excel 已保存的 cached error。

## 證據三：Calc 重算檔保留連結，但沒有外部資料快取

Calc 輸出仍保留 `externalLink1.xml`，且仍列出外部工作表 `BlockList`，但 `<sheetData>` 已變成空的：

```xml
<sheetData sheetId="0" />
```

其 relationship 也仍指向同一個不存在於目前環境的外部工作簿，只是路徑被 Calc 改寫成：

```text
<original workstation path>/SM2734_HWS_SA_FMEDA_0.2chk.xlsx
```

這個現象表示 Calc 沒有把外部工作簿資料嵌入或成功刷新；它不是拿到一份完整的外部 `BlockList` 後算出 0。

## 證據四：T2/T3 變成 0，W2/W3 因算術傳播變成 0

在原始工作簿中，`U2=172`、`V2=75`，`U3=173`、`V3=306`，這些本地輸入都是正常數字。下游公式為：

```excel
W2=(172+75)*T2/1024/1024
W3=(173+306)*T3/1024/1024
```

因此只要 Calc 把外部 `SUMIF` 的 `T2`、`T3` 視為 0，就會得到：

```text
W2 = 247 × 0 / 1024 / 1024 = 0
W3 = 479 × 0 / 1024 / 1024 = 0
```

這完整解釋了為什麼 4 個 error-state changes 會同時出現：**T2/T3 是根因，W2/W3 只是錯誤傳播被 0 取代後的下游結果。**

另外，`X2` 與 `X3` 並未變成正常值，而是從原始 `#VALUE!` 變成 Calc 的 `#DIV/0!`；這是同一個外部資料缺失的旁證，因為其公式仍然引用 `[1]BlockList!C3`／`[1]BlockList!C2`。

## 排除其他可能原因

| 候選原因 | 判斷 | 信心 | 理由 |
|---|---|---|---|
| 公式被轉換或改寫 | 非主要原因 | 高 | 四格公式 XML 文字完全相同；公式 blocking changes 在全量 validation 為 0 |
| 本地 `U`／`V` 輸入錯誤 | 非主要原因 | 高 | U2、V2、U3、V3 都是正常數字且原始／重算一致 |
| 合併儲存格造成公式錯位 | 非主要原因 | 高 | `SRAM Tran FIT` 的 6 個 merged ranges 在 validation 中 preserved=True |
| 浮點精度誤差 | 不適用 | 高 | 這 4 格是錯誤字串到數字 0，不是極小數值差異 |
| Python parser 造成錯誤 | 非主要原因 | 高 | 直接比較 XLSX XML；原始檔已保存 `t="e"` 與 `#VALUE!` |
| 外部工作簿未提供／Calc 無法存取 | **最主要原因** | 很高 | 公式使用 `[1]BlockList`、原始 `refreshError=1`、重算檔外部 sheetData 空、目前環境找不到目標檔 |
| Calc 與 Excel 對未解析外部連結的處理不同 | 直接原因之一 | 高 | 同一公式原始 cache 是 `#VALUE!`，Calc cache 是 0，且外部快取仍為空 |

## 是否可以接受這 4 格變成 ok？

目前答案是：**不可以直接接受。** 原因不是「Calc 一定算錯」，而是 0 的來源沒有被證明是外部 `BlockList` 真實計算所得。它可能代表空資料、未解析外部連結，或 Calc 對缺失外部來源的 fallback 行為。

對 FMEDA 而言，`#VALUE! → 0` 不是一般的顯示差異。它可能把「沒有資料／沒有成功更新」誤讀成「計算結果確實為零」，而這會影響後續的 failure rate、FIT、診斷覆蓋率或安全評估結果。因此目前應保留為：

```text
status = unresolved_external_recalculation
acceptance = blocked
reason = external workbook unavailable or not refreshed
```

`W2` 與 `W3` 也應連帶標示為 `derived_from_unresolved_external`，不能只記錄它們現在的數值為 0。

## 建議的修正順序

### 第一優先：補齊並驗證外部工作簿

取得並放置原始公式指向的：

```text
SM2734_HWS_SA_FMEDA_0.2chk.xlsx
```

接著在同一個可控環境中重新執行 Calc，確認：

1. external link 能成功開啟。
2. `externalLink1.xml` 不再有 refresh error。
3. 外部 `BlockList` 的 `R10:R49` 與 `AO10:AO49` 實際可讀。
4. `T2`／`T3` 的結果不是由空資料 fallback 產生。
5. `W2`／`W3` 與 `X2`／`X3` 的錯誤狀態是否同步消失。

### 第二優先：把外部依賴納入 workspace manifest

目前 workspace 不應只保存外部引用字串，還應增加：

| 欄位 | 用途 |
|---|---|
| `external_workbook_original_uri` | 保存 Excel 原始外部路徑 |
| `external_workbook_resolved_path` | 保存這次實際解析到的檔案 |
| `external_workbook_sha256` | 確認外部檔案版本 |
| `external_sheet` | 例如 `BlockList` |
| `refresh_status` | `resolved`／`unresolved`／`refresh_error` |
| `dependent_cells` | 例如 T2、T3、W2、W3、X2、X3 |
| `acceptance` | `blocked`／`reviewed`／`accepted` |

### 第三優先：讓 validator 具備錯誤語意傳播

目前 validator 已經能判斷 error-state change，但下一步應讓報告明確說明：

```text
T2/T3 = root external dependency changes
W2/W3 = downstream derived changes
X2/X3 = remaining external denominator failures
```

如此審查者不必從 2,784 筆差異中自己找出因果鏈。

### 第四優先：增加人工審查 gate，而不是放寬 PASS

在外部工作簿補齊前，建議保持 `FAIL`。補齊外部來源後，只有在以下證據都成立時，才允許改為 `PASS_WITH_WARNINGS` 或 `PASS`：

- 外部工作簿 hash 已記錄。
- refresh status 已成功。
- T2/T3 的 SUMIF 結果可回查外部資料列。
- W2/W3 的 0 或非零結果能由 T2/T3 重算推導。
- X2/X3 的錯誤狀態有明確解釋。
- FMEDA 工程師對抽樣結果完成確認。

## 最終判定

這 4 格應歸類為：

```text
Root cause category: unresolved external workbook / recalculation engine boundary
Severity: blocking for automatic acceptance
Confidence: high
Current acceptance: reject automatic PASS
Required next action: provide and resolve SM2734_HWS_SA_FMEDA_0.2chk.xlsx, then rerun
```

這次分析也驗證了一個重要設計原則：**error-state change 不能只看「現在是不是錯誤」，還要追蹤錯誤是如何消失的、它是否只是被外部資料缺失轉換成 0，以及下游結果是否因此被錯誤地美化。**
