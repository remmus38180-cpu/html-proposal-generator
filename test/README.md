# 測試

需要 Node 22+（用到內建 `CompressionStream` / `Blob`）與 `playwright`。

```bash
npm i playwright                      # 僅 ui.mjs 需要
node test/mkfix.mjs                   # 產生 test/fixtures/ 的假主檔與案件檔
node test/t3.mjs                      # [C11]/[C12]/[D19] 截斷與金額格式邊界
node test/smoke.mjs                   # 模組載入 + [A1]~[D19] 規則 + 四種產出（無瀏覽器）
CHROME=... node test/ui.mjs           # 端對端（分檔版）：上傳 xlsx、選路徑、按按鈕、抓 console error 與下載
node build/bundle.mjs                 # 產生 dist/case_template.html 單檔版
CHROME=... node test/ui-bundle.mjs    # 端對端（單檔版）：複製到空目錄後以 file:// 開啟
```

`ui.mjs` 會在 8731 埠起一個靜態伺服器指向專案根目錄，並驗證：

- 初次載入**只有**路徑 A/B 共用的 7 個 js
- price 年月欄位不存在時不解鎖路徑選擇
- 選路徑 A 後**不會**載入 price-calc / case-report / docx-writer / docx-report
- 選路徑 B 才動態注入那 6 個模組
- B → A 切換跳確認對話框，且 fCase 狀態被清掉
- 四種產出（查詢表／Word／四分頁／附件）都能下載且為合法 ZIP

`ui-bundle.mjs` 會把 `dist/case_template.html` 複製到一個**只有這一個檔**的臨時目錄，
以 `file://` 開啟後驗證：

- 沒有任何額外的 `.js` 網路請求（全部內嵌）
- 選路徑 A 時 `buildCaseReport` / `costRefPrice` / `reportToDocx` 皆為 `undefined`
  —— 單檔版同樣保有「路徑 A 不執行試算與 Word 模組」的隔離
- 選路徑 B 後才變成 `function`
- 五種產出（查詢表／Word／四分頁／附件／範本）都能下載

`CHROME` 未設時用 playwright 自帶的瀏覽器。
唯一預期會出現的 console error 是 Google Fonts CDN 連不到（離線環境），字型會自動降階。
