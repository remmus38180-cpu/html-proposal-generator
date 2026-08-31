# html-proposal-generator
一個純前端 HTML 提案產生器，能解析使用者條件並動態組裝產出客製化的 DOCX、XLSX 及 PPTX 檔案 (A client-side HTML proposal generator that parses user conditions to dynamically build and export tailored DOCX, XLSX, and future PPTX files. )
proposal-studio-web/
├── index.html               <-- 主介面與應用程式入口
├── README.md
├── assets/
│   ├── js/
│   │   ├── app.js           <-- 條件判斷與邏輯
│   │   └── exporters/       <-- .xlsx / .docx / .pptx 匯出模組
│   └── css/
└── templates/               <-- 提案條件條文與 Layout 結構
