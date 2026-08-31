# html-proposal-generator

一個純前端 HTML 提案產生器，能解析使用者條件並動態組裝產出客製化的 DOCX、XLSX 及 PPTX 檔案 (A client-side HTML proposal generator that parses user conditions to dynamically build and export tailored DOCX, XLSX, and future PPTX files.)

## 專案架構

```text
html-proposal-generator/
├── index.html                    # 純前端 UI 入口 (支援路徑 A / 路徑 B 雙軌切換)
├── docs/                         # 系統與輸入資料規格文件
├── legacy/                       # 舊版單軌程式碼 (備份與重構比對用)
└── assets/
    └── js/
        ├── app.js                # 主控邏輯與路徑狀態 (TRACK_A / TRACK_B) 切換
        ├── parser.js             # 數據清洗 (浮點數 roundFloat 與空值修整)
        ├── group-model.js        # [共用] 分頁 2~4 歷史申報/異動/十國藥價模型
        ├── price-calc.js         # [路徑 B 獨有] 分頁 1 藥價核定與財務試算
        └── exporters/            # SheetJS (.xlsx) 與 docx.js (.docx) 導出模組
```
