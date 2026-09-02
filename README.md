# html-proposal-generator

健保藥品調價試算產生器。純前端、零外部函式庫，全部運算在瀏覽器內完成，不上傳任何資料。
**雙軌分流**：只有藥品主檔時走路徑 A 做背景查詢；備妥案件需求檔後走路徑 B 產出正式提案。

## 使用方式

用瀏覽器直接開 `index.html`（本機檔即可，`assets/` 需與它同層）。

| | 路徑 A｜背景查詢 | 路徑 B｜正式提案 |
|---|---|---|
| 輸入 | 藥品主檔 ＋ price 民國年月 | 再加上案件需求檔 |
| 輸出 | 查詢表 Excel（3 個工作表） | 提案 Word ＋ 四分頁 Excel ＋ 提案附件 Excel |
| 用途 | 廠商剛送件，釐清藥品身份、同分組相似品、市場規模 | 完整核價試算與公文產生 |

操作流程：上傳主檔 → 系統掃描 `PRICE{年月}` 欄位並列出可選年月 → 選定年月（驗證欄位存在後才解鎖）
→ 選路徑 → （路徑 B 才出現案件需求檔上傳區）→ 產出。

## 專案架構

```text
html-proposal-generator/
├── index.html                        # 雙軌 UI 入口
├── INTENT.md                         # 雙軌重構規格
├── claude.md                         # 系統規格（權威來源）
├── 藥品主檔資料處理原則.md            # 法規與資料處理規則 [A1]~[D20]（權威來源）
├── 資料邏輯說明.html                  # 給主管／同仁的說明頁
├── legacy/
│   └── case_template.html            # 重構前的單軌單檔版，凍結供比對
└── assets/js/
    ├── format-utils.js               # [共用] num/txt/ad8/truncatePrice/chineseAmount…
    ├── parser.js                     # [共用] ZIP 讀寫 + XLSX / CSV 解析
    ├── xlsx-writer.js                # [共用] XLSX 組出引擎
    ├── group-model.js                # [共用] 分組模型核心：[A2][A3][B5]~[B8]
    ├── app.js                        # [共用] DUAL_TRACK 狀態與路徑分流
    ├── price-calc.js                 # [路徑 B] 核價試算 [C9]~[C14]
    ├── case-report.js                # [路徑 B] 案件需求檔 → 提案報表模型
    ├── docx-writer.js                # [路徑 B] DOCX OOXML 引擎
    ├── template.js                   # [路徑 B] 案件需求檔範本
    └── exporters/
        ├── sas-workbook.js           # [共用] 四分頁報表（分頁 1 以 skipPriceCalc 控制）
        ├── query-workbook.js         # [路徑 A] 背景查詢表 3 分頁
        ├── attachment-workbook.js    # [路徑 B] 提案附件
        └── docx-report.js            # [路徑 B] 提案 Word 組字
```

### 模組載入策略

`index.html` 只用 `<script src>` 靜態載入路徑 A/B 共用的 6 個模組。
標記為「路徑 B」的 6 個模組在使用者**選擇路徑 B 時才動態注入**，
因此路徑 A 的頁面上不存在 `price-calc.js` / `case-report.js` / `docx-writer.js` /
`docx-report.js`，無法誤觸試算邏輯。

模組之間以全域函式互相呼叫（無打包步驟、無 ES module），載入順序即相依順序。

## 規則對照

計算規則以 `藥品主檔資料處理原則.md` 為準，實作位置：

| 規則 | 內容 | 實作 |
|---|---|---|
| [A1] | 年度視窗（`最新申報年度資料範圍` 判定完整 12 個月年度） | `group-model.js` `resolveYearWindow` / `parseDataRange` |
| [A2] | QTY/AMT 採計四情形（含不良品暫停支付註記） | `group-model.js` `buildGroupModel` |
| [A3] | 預設排除 TPN／不上網 | `group-model.js` `excludeMaster` |
| [B5] | 同分組項目數 | `group-model.js` `countValid` / `isCounted` |
| [B6] | 近三年平均申報量（分母固定 3/2/1） | `group-model.js` `buildGroupModel` |
| [B7] | 同分組每月申報金額 | `group-model.js` `buildGroupModel` |
| [B8] | 同成分同劑型收載年 + 第三 B 大類 | `group-model.js` `ingredientKey` / `buildGroupModel` |
| [C9]~[C14] | 加成級距、領證加計、浮點防護、第 40 條 | `price-calc.js`、`format-utils.js` |
| [D15]~[D20] | 分類對照、金額級距、占率位數 | `format-utils.js` |
