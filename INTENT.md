# 健保藥品調價試算產生器：雙軌分流架構重構

## 背景

單軌架構痛點：
- 廠商案件初期，使用者無案件需求檔（fCase），但需快速釐清藥品基礎資訊（ATC、分組、同分組品項、收載歷史）
- 目前只能等備妥 fCase 才能進工具，無法分階段作業

## 目標

**主目標**：架構從單軌（fMaster + fCase → 提案文件）改為雙軌分流

**路徑 A｜藥品背景查詢**
- 輸入：藥品主檔（fMaster） + 設定 price 民國年月
- 輸出：查詢表（3 個 Excel 工作表）
- 內容：CODE、ATC7、分組名稱/代碼、同分組項目、收載日期、申報量趨勢
- 用途：廠商剛送件，使用者查詢藥品身份、相似品、市場規模

**路徑 B｜正式提案製作**（既有功能）
- 輸入：fMaster + fCase
- 輸出：Word 提案 + Excel 核價試算（4 張表）+ 提案附件 Excel
- 用途：完整核價試算與公文產生

---

## 技術層面：代碼現況與必修項

### 已發現的 Bug（優先修）

#### [A2] buildGroupModel 年度歸零漏檢不良品註記
**位置**：case_template.html 第 1206~1209 行

**現況**：
```js
if (p === 0 && eff !== null){
  for (j=0;j<3;j++) if (!(yrs[j] <= eff)){ qa[j] = 0; aa[j] = 0; }
}
```

**問題**：僅檢查「支付價=0 且有生效日期」就歸零晚於生效年的 QTY/AMT，**完全無檢查不良品暫停支付註記**

**法規依據**（藥品主檔資料處理原則 [A2]）：
- 情形 1：不良品暫停支付註記 = "Y" → QTY/AMT 直接採計，**無年份邊界限制**
- 情形 3：支付價 = 0 + 不良品註記 ≠ "Y" → 始得歸零晚於生效年的年度

**影響範圍**：
- B6（近三年平均申報量）→ 錯算
- B7（同分組每月申報金額）→ 錯算
- C9（加成級距判定）→ 誤判
- Path A、Path B 輸出皆受影響

**修正**：判斷式改為 `if (p === 0 && susp !== 'Y' && eff !== null)`

---

#### [B8] 同成分同劑型收載年判定【完全缺失】

**法規依據**（藥品主檔資料處理原則 [B8]）：
- 用途：判斷是否超過 15 年（第三 B 大類標記）
- 基準年：依 PRICE 年月決定
  - 月份 ≥ 4 月：基準年 = PRICE 民國年
  - 月份 < 4 月：基準年 = PRICE 民國年 − 1
- 同成分同劑型識別：分組代碼前 8 碼 + 第 10、11 碼（核價劑型碼）
- 判定：基準年 − 同成分同劑型首次收載年 ≥ 15 → 第三 B 大類

**現況**：程式碼中完全無相關邏輯

**需補**：
1. 建立 Map，key = 分組代碼前 8 碼 + 第 10、11 碼，value = MIN(收載日期) 民國年
2. 根據 price 年月計算基準年
3. 每筆記錄判斷並標記 `thirdBLarge` 欄位

**優先級**：Path B 試算會用（Path A 可不顯示，但 model 要正確）

---

### 代碼結構拆分方案

| 模組 | 內容 | 路徑 A | 路徑 B | 說明 |
|------|------|:----:|:----:|------|
| **parser.js** | Zip/XLSX/CSV 讀寫、sheet 解析 | ✅ | ✅ | 基礎 I/O |
| **format-utils.js** | num/fmt/money/chineseAmount/ad8/catLabel/adaptivePct | ✅ | ✅ | 純格式函式 |
| **group-model.js** | buildGroupModel + excludeMaster + [B8]判定 | ✅ | ✅ | **共用核心，[A2] 改在這** |
| **price-calc.js** | buildPriceCalc/priceCalcHeaders/a10RefPrice/costRefPrice | ❌ | ✅ | 路徑 A 完全不載 |
| **case-report.js** | mergeCaseRowsByCode/buildCaseReport | ❌ | ✅ | 路徑 A 完全不載 |
| **xlsx-writer.js** | XLSX 組出引擎（styles/sheet/buildXlsxBook） | ✅ | ✅ | 共用 |
| **exporters/sas-workbook.js** | buildSasWorkbook（分頁 2/3/4 共用 + 分頁 1 改選用） | ✅ | ✅ | **分頁 1 用參數 skipPriceCalc 控制** |
| **exporters/attachment-workbook.js** | buildAttachmentWorkbook（格式依路徑微調） | ✅ | ✅ | 共用，Path A 版略簡 |
| **docx-writer.js** | OOXML 段落/表格/編號引擎（底層） | ❌ | ✅ | 路徑 A 完全不載 |
| **exporters/docx-report.js** | reportToDocx（Word 提案組字） | ❌ | ✅ | 路徑 A 完全不載 |
| **template.js** | 案件需求檔範本 | ❌ | ✅ | 路徑 A 完全不載 |
| **app.js** | UI 流程、路徑分流、權限檢查、getModel/refresh | ✅ | ✅ | **核心改寫** |

**架構約束**：
- Path A HTML 頁面**不引入** price-calc.js / case-report.js / docx-writer.js / docx-report.js，強制防止誤觸試算邏輯
- Path B HTML 頁面正常引入全部

---

## UI 流程與狀態管理

### 操作步驟

```
[1] 上傳藥品主檔（fMaster）
    ↓
[2] 系統掃描 PRICE{年月} 欄位 → 列出可用年月清單
    ↓
[3] 使用者選年月（如 11508）
    ↓
[4] 系統驗證欄位存在 → 執行 buildGroupModel(11508)
    ↓
[5] 顯示路徑選擇（Radio button）
    ⭕ 背景查詢
    ⭕ 正式提案
    ↓
路徑 A → 執行 buildGroupModel，產出 Sheet 2/3/4（跳過核價試算）
路徑 B → 出現 fCase 上傳區 → mergeCaseRowsByCode → 完整試算
```

### 全域狀態結構

```javascript
const DUAL_TRACK = {
  path: null,                    // 'PATH_A' | 'PATH_B'
  priceYearMonth: null,          // '11508'
  
  // 共用層（Path A/B 都用）
  masterData: {
    parsed: null,                // xlsx 原始解析結果
    model: null,                 // buildGroupModel(year) 結果，含 [B6/B7/B8]
    exclusions: { tpn: [], offnet: [] }
  },
  
  // 路徑 A
  pathA: {
    sheets: { sheet2: null, sheet3: null, sheet4: null }
  },
  
  // 路徑 B
  pathB: {
    caseData: null,              // mergeCaseRowsByCode 結果
    targets: null,               // buildCaseReport 結果（含試算）
    outputs: {
      word: null,                // Word 提案
      xlsxMain: null,            // 4 分頁報表
      xlsxAttach: null           // 提案附件
    }
  }
};
```

### 清空規則

| 事件 | 清空內容 | 保留內容 |
|------|--------|--------|
| 上傳新 fMaster | DUAL_TRACK 全部 | 無 |
| 改選 price 年月 | masterData.model / pathA / pathB | 檔案本身 |
| 切換路徑 A → B | pathA.sheets | masterData |
| 切換路徑 B → A | pathB.caseData / pathB.outputs | masterData + 詢問確認 |
| 上傳新 fCase | pathB.outputs 只 | masterData + pathB.caseData |

---

## 輸出格式差異

### 路徑 A（背景查詢版）

**Excel 結構**：3 個工作表

| 工作表 | 內容 | 欄位 |
|------|------|------|
| **Sheet 1** | 查詢摘要 | CODE / 藥品名 / 廠商 / ATC7 / 分組代碼 / 分組名稱 / 收載日期 / 同分組項目數 |
| **Sheet 2** | 同分組比較 | 分組內全部藥品 CODE / 藥品名 / 廠商 / ATC7 / 收載日期 / 近三年平均申報量 / 申報金額 |
| **Sheet 3** | 同 ATC 相關品項 | 同 ATC7 碼的全部品項（可能不同分組） |

**簡化規則**：
- ❌ 無 PRICE、廠商建議價、十國價、成本價（試算用）
- ❌ 無調高後支付價、財務衝擊（試算用）
- ❌ 無核定備註（提案用）
- ✅ 保留分組層彙總（支付價區間、項目數、申報量）
- 頁籤名：`背景查詢_{CODE}` 或 `查詢 {藥品名}`
- 頁首/頁尾加註：「此為背景查詢版，非正式提案 / 查詢日期 {YYYY-MM-DD}」

---

### 路徑 B（正式提案版）

**維持現狀**：Word + 4 分頁 Excel + 提案附件 Excel

**修改**：
- buildSasWorkbook 分頁 1 改為可選參數
  ```js
  buildSasWorkbook(model, targets, { skipPriceCalc: false })
  ```

---

## 錯誤處理與驗證順序

| 檢查點 | 位置 | 提示 | 恢復 |
|------|------|------|------|
| **fMaster 缺必要欄** | 上傳時（嚴格） | ❌ 缺少 CODE / 分組代碼 / ATC7 / PRICE 等 | 重新上傳 |
| **可用年月掃描** | 解析完成 | ✅ 偵測 PRICE 欄位清單 | 使用者選擇 |
| **price 年月無對應欄** | 選年月時 | ❌ PRICE11508 不存在，可選：11507、11509 | 改選年月 |
| **price 年月驗證通過** | 確認後 | ✅ buildGroupModel 執行成功 | 解鎖路徑選擇 |
| **路徑 B 無 fCase** | 點執行前 | ⚠️ 已選正式提案，請上傳案件檔 | 上傳 fCase |
| **fCase 與 fMaster 無匹配** | 解析 fCase | ⚠️ CODE x, y 未在主檔（移除該列） | 檢查或接受 |
| **groupModel 內部異常** | buildGroupModel 中 | ❌ 某列資料無法處理（詳見診斷資訊） | 檢視 master 原始資料 |

---

## 成功標準

✅ **功能完成**
- Path A：上傳 fMaster + 選 price 年月 → 下載查詢表（3 張 Excel）
- Path B：上傳 fMaster + fCase + 選 price 年月 → 下載 Word + Excel（4 張）+ 提案附件
- 路徑切換時，狀態正確隔離（不會混淆）

✅ **Bug 修正**
- [A2] 年度歸零考慮不良品註記 → 近三年平均申報量正確
- [B8] 同成分同劑型收載年判定 → 第三 B 大類標記正確
- buildSasWorkbook 分頁 1 參數化 → 無歧義

✅ **架構品質**
- price-calc.js / docx-writer.js / docx-report.js 在 Path A 頁面**不被載入**
- group-model.js 修正後，Path A/B 用同一份正確的 model
- 模組分界清楚，便於後續維護

✅ **使用者體驗**
- 無冗餘欄位（Path A 不顯示試算項）
- 錯誤提示清晰、可恢復
- 路徑選擇直觀（Radio button，同一頁）

---

## 實作建議

### 第 1 階段：修正 + 模組化
1. 修正 [A2] bug（buildGroupModel）
2. 補充 [B8] 邏輯（group-model.js）
3. 拆分 app.js → 引入 DUAL_TRACK 狀態管理
4. buildSasWorkbook 分頁 1 參數化

### 第 2 階段：UI 流程
1. 上傳 fMaster + price 年月驗證
2. 路徑選擇界面（Radio button）
3. 路徑 B 的 fCase 上傳區（動態顯示）

### 第 3 階段：導出
1. Path A 導出（Sheet 2/3/4）
2. Path B 導出（維持現狀，確保與 Path A 輸出無串聯）

---

## 限制與假設

- **fMaster 欄位動態**：支援 PRICE{年月} 動態識別，不硬編碼
- **price 年月必選**：必須明確指定，不自動推斷
- **Path A 查詢版無試算**：即使分組有 0 元品項也只顯示「−」，不計算調高後支付價
- **localStorage 暫不做**：10 分鐘暫存 fMaster 可後期補，MVP 先無
- **多年月對比暫不做**：Path A 一次只查一個年月

---

## 工作量估計

| 項目 | 工作量 |
|------|------|
| [A2] / [B8] bug 修正 + 測試 | 中 |
| group-model.js 重構 + 模組拆分 | 中 |
| app.js DUAL_TRACK 改寫 + 權限檢查 | 中 |
| HTML UI（price 年月 + 路徑選擇） | 小 |
| exporters 調整（sas/attachment 格式微調） | 小 |
| **總計** | **中~大** |

---

## 檢查清單（Opus 可用）

### 法規符合性（對照「藥品主檔資料處理原則.md」）

- [ ] [A1] 年度視窗判定：支付價欄位對應民國年；近三年 = 完整 12 個月年份 + 前 2 年
- [ ] [A2] 採計規則四情形：
  - [ ] 情形 1（不良品註記=Y）：QTY/AMT 直接採計，無邊界限制
  - [ ] 情形 2（不良品註記≠Y 且 PRICE≠0）：直接採計
  - [ ] 情形 3（不良品註記≠Y 且 PRICE=0）：晚於生效年的年度歸零 ← **[A2] bug 修正點**
  - [ ] 情形 4（PRICE 為空值）：直接採計
- [ ] [A3] 預設排除：TPN 前 3 碼 + 不上網註記，在 buildGroupModel 最前執行
- [ ] [B5] 項目數計算：支付價≠0 或（支付價=0 且不良品註記=Y）才計入
- [ ] [B6] 近三年平均申報量：
  - [ ] 分組收載年 = 該分組所有項目收載日期的最小值
  - [ ] 早於分組收載年的年度不納入（分母不因 NULL 縮減）
- [ ] [B7] 同分組每月申報金額 = 當期年度 AMT 總計 ÷ 12（加成級距判定依據）
- [ ] [B8] 同成分同劑型收載年判定（**完全缺失，必補**）：
  - [ ] 識別碼 = 分組代碼前 8 碼 + 第 10、11 碼
  - [ ] 基準年判定：月份≥4 月用當年，<4 月用前年
  - [ ] 判定邏輯：基準年 − 首次收載年 ≥ 15 → 第三 B 大類
- [ ] [C9] 加成級距（支付價區間或成本價區間）正確應用於核價
- [ ] [C10] 領證加計：領有許可證(Y) → ×1.0505（營業稅 5% + 藥害救濟 0.05%）
- [ ] [C11] 浮點誤差防護：toPrecision(10) 清誤差後再截斷
- [ ] [C12] 第 40 條進位規則：<5 兩位 / 5~50 一位 / ≥50 整數
- [ ] [D15] 藥品分類對照：代碼 1~6, 9 與分類名（研發廠/生物製劑/BA-BE/學名藥/相似性/對照品）
- [ ] [D16] 大類對照：分類 1~6 與第一/二/三A/三B 大類
- [ ] [D17] 必要藥品代碼：0/1/3/4 與特殊/一般/不可替代標記
- [ ] [D18] ATC_MODE 三模式：FULL(7碼全等) / PREFIX5(前5碼) / BOTH(預設，兩者都做) → 表二產出張數
- [ ] [D19] 金額顯示級距：元/萬/億/兆 自動轉換（chineseAmount2 與 chineseAmount 展開式）
- [ ] [D20] 占率位數：10% 以上整數 / 1%~10% 一位小數 / <1% 補到首個有效數字

### 功能與架構

- [ ] [A2] bug 修正（年度歸零加入不良品檢查）
- [ ] [B8] 邏輯補充（同成分同劑型收載年 + 15 年判定）
- [ ] group-model.js 修正後驗證（Path A/B 輸出數字一致）
- [ ] DUAL_TRACK 狀態機運作正常（路徑切換、清空規則符合預期）
- [ ] Path A 頁面**不載入** price-calc/docx-writer/docx-report 模組
- [ ] buildSasWorkbook 分頁 1 `skipPriceCalc` 參數測試
- [ ] Path A 查詢表格式無試算欄
- [ ] 錯誤攔截工作（fMaster 缺欄、price 年月無對應、fCase 無匹配 CODE）

### 分組代碼與輸出格式

- [ ] 分組代碼編碼原則遵守（第 684~750 行）：
  - [ ] 代碼結構正確（12 碼：ATC3+成分3+流水1+成分數1+含量序1+劑型2+規格1）
  - [ ] 同成分同劑型識別：前 8 碼+第 10、11 碼
  - [ ] 罕藥證分組名稱後標註「(罕)」
- [ ] 附錄 A 網頁字體規範（第 755~782 行）：
  - [ ] HTML 引入 Google Fonts（IBM Plex Mono / Noto Sans TC）
  - [ ] 預設字族：`'IBM Plex Mono', 'Noto Sans Mono CJK TC', 'Noto Sans TC', 'PingFang TC', 'Microsoft JhengHei', monospace`
  - [ ] 表格/code 區塊加 `font-variant-numeric: tabular-nums;`
- [ ] 附錄 B claude-light.css 規範（第 785~833 行）：
  - [ ] Word 提案輸出遵守色彩規格（背景白、正文深灰、code 紅底淺灰）
  - [ ] 尺寸：最大寬 860px、padding 48px、行高 1.7
  - [ ] CSS 內嵌（不外連），確保 PDF 列印完整

### 測試驗證

- [ ] 可靠性測試：同一份資料用舊程式 vs 新程式，Path B 輸出應 100% 一致
- [ ] Path A/B 數字一致性：分組層金額/數量應相同（Path A 無核價欄，但 model 中的申報量要一致）
- [ ] Path A 查詢欄位完整性：CODE、ATC7、分組代碼/名稱、同分組項目、收載日期、近三年申報量
- [ ] Path B Word 提案驗證：
  - [ ] 表一/表二/表五/表六 格式正確
  - [ ] ATC_MODE = FULL 時一張表、BOTH 時兩張表
  - [ ] 法條文字、加計比例欄位正確對應加成級距
- [ ] Path B Excel 產出驗證：
  - [ ] 分頁 1（藥價核定試算）：targets 非空時含試算、空時含提示
  - [ ] 分頁 2/3/4 數字與 Word 表二/表五一致
- [ ] 邊界情況：
  - [ ] 分組無收載日期
  - [ ] 全部品項支付價=0
  - [ ] 不良品註記=Y 且支付價=0
  - [ ] price 年月晚於最新申報年份
  - [ ] ATC_MODE=FULL 與 PREFIX5 的表產出張數
