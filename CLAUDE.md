# html-proposal-generator — CLAUDE.md

專案概述與 AI 協作紀錄。

---

## 專案架構

- **雙軌分流**：路徑 A（背景查詢，不試算）/ 路徑 B（正式提案，含核價試算）
- 原始碼位於 `assets/js/`；`node build/bundle.mjs` 輸出 `dist/case_template.html`（單檔版）
- 路徑 B 模組以 `<script type="text/plain" data-mod="...">` 惰性內嵌，選擇路徑 B 時才由 `injectModule()` 執行

---

## 修訂紀錄

### 2026-09-04　路徑 A 查詢表重構（branch: fix/a2-suspend-b8-ingredient-year）

**異動檔案**
- `assets/js/exporters/query-workbook.js`
- `assets/js/app.js`（btnQuery log message）

**Sheet 結構變更：3 表 → 2 表**

| 原工作表 | 新工作表 | 說明 |
|---|---|---|
| 1. 查詢摘要（直接命中篩選條件的品項） | 廢除 | 合併進新 Sheet 1，以底色標記 |
| 2. 同分組比較（isCounted 品項） | **1. 同分組項目明細** | 改為展開分組內**全部品項**（含 0 元），新增「含量」「規格量」欄，命中篩選條件的品項以 MARK 色標記 |
| 3. 同ATC相關品項 | **2. 分組與ATC彙總** | 改為 分組代碼×ATC7碼 聚合層級（非品項層級） |

**Sheet 1 同分組項目明細**
- 先依篩選條件定位目標分組代碼，再展開該分組旗下**所有品項**（含 0 元）
- 新增欄位：`含量`（`成分含量` || `成分及含量`）、`規格量`（`規格量`）
- 命中篩選條件的品項以 `XS.MARK`（淡藍底色）標記藥品名稱欄
- 排序：調劑大類↑ → 分組代碼↑ → 0元置後 → 許可證層AMT↓ → 當期AMT↓ → CODE↑
- 每個分組結束插入灰底小計列；末尾加總計列

**Sheet 2 分組與ATC彙總**
- 聚合鍵：`分組代碼 × ATC7碼`
- 遵守計數規則：排除「支付價＝0 且不良品暫停支付註記≠Y」品項（`isCounted` 過濾）
- 項目數、申報量/金額均只計算 `isCounted` 通過的品項
- 若某分組只有一種 ATC，**不輸出該分組的小計列**
- 排序：分組代碼↑ → ATC7碼↑

**counts 物件變更**

```js
// 舊
{ sel, groups, s1, s2, s3, atc }
// 新
{ sel, groups, s1, s2 }
```

---

## 資料處理原則（路徑 A）

- `isCounted(r)`：`price !== 0` 或 (`price === 0` 且 `suspFlag(r.d) === true`)
- `qtyAdj` / `amtAdj`：0 元品項在生效日後歸零（已由 group-model 處理）
- Sheet 1 明細列出全部品項（不套 isCounted），讓使用者看到分組全貌
- Sheet 2 彙總計數套 isCounted，與正式提案統計口徑一致
