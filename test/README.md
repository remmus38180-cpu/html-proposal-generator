# 測試

需要 Node 22+（用到內建 `CompressionStream` / `Blob`）與 `playwright`。

```bash
npm i playwright                      # 僅 ui.mjs 需要
node test/mkfix.mjs                   # 產生 test/fixtures/ 的假主檔與案件檔
node test/t3.mjs                      # [C11]/[C12]/[D19] 截斷與金額格式邊界
node test/smoke.mjs                   # 模組載入 + [A1]~[D19] 規則 + 四種產出（無瀏覽器）
CHROME=/opt/pw-browsers/chromium-1194/chrome-linux/chrome node test/ui.mjs
                                      # 端對端：真的上傳 xlsx、選路徑、按按鈕、抓 console error 與下載
```

`ui.mjs` 會在 8731 埠起一個靜態伺服器指向專案根目錄，並驗證：

- 初次載入**只有**路徑 A/B 共用的 7 個 js
- price 年月欄位不存在時不解鎖路徑選擇
- 選路徑 A 後**不會**載入 price-calc / case-report / docx-writer / docx-report
- 選路徑 B 才動態注入那 6 個模組
- B → A 切換跳確認對話框，且 fCase 狀態被清掉
- 四種產出（查詢表／Word／四分頁／附件）都能下載且為合法 ZIP

`CHROME` 未設時用 playwright 自帶的瀏覽器。
