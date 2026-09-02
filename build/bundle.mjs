/* 把 index.html + assets/js/** 併成一個可寄送的單檔 HTML。
 *
 *   node build/bundle.mjs
 *   → dist/case_template.html
 *
 * 兩種形態的行為完全相同：
 *   - 路徑 A/B 共用模組 → 直接內嵌為 <script>，載入時執行
 *   - 路徑 B 專屬模組   → 內嵌為 <script type="text/plain" data-mod="…">，
 *                          由 app.js 的 injectModule() 在選擇路徑 B 時才執行
 * 因此單檔版同樣保有「路徑 A 不執行試算與 Word 模組」的隔離。
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const ROOT = process.env.APP_ROOT || path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT_DIR = path.join(ROOT, 'dist');
const OUT = path.join(OUT_DIR, 'case_template.html');

/* <script type="text/plain"> 區塊會被任何 </script 字串提前截斷 */
function safeForInline(js){
  return js.replace(/<\/script/gi, '<\\/script');
}
function read(rel){
  const p = path.join(ROOT, rel);
  if (!fs.existsSync(p)) throw new Error('找不到模組：' + rel);
  return fs.readFileSync(p, 'utf8').replace(/\s*$/, '\n');
}

let html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');

/* 1. app.js 裡宣告的路徑 B 模組清單，就是要延後執行的那幾支 */
const appSrc = read('assets/js/app.js');
const m = appSrc.match(/var PATH_B_MODULES = \[([\s\S]*?)\];/);
if (!m) throw new Error('app.js 找不到 PATH_B_MODULES 宣告');
const pathB = [...m[1].matchAll(/'([^']+)'/g)].map(x => x[1]);

/* 2. index.html 裡的 <script src> 就是靜態載入的共用模組 */
const staticSrcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map(x => x[1]);
if (!staticSrcs.length) throw new Error('index.html 找不到任何 <script src>');

const blocks = [];
blocks.push('<!-- ===== 路徑 A / B 共用模組（內嵌，載入即執行） ===== -->');
for (const src of staticSrcs){
  blocks.push('<!-- ' + src + ' -->');
  blocks.push('<script>\n' + safeForInline(read(src)) + '</script>');
}
blocks.push('');
blocks.push('<!-- ===== 路徑 B 專屬模組（內嵌但不執行，選擇路徑 B 時才由 injectModule() 執行） ===== -->');
for (const src of pathB){
  blocks.push('<script type="text/plain" data-mod="' + src + '">\n' + safeForInline(read(src)) + '</script>');
}

/* 3. 換掉原本那串 <script src>；路徑 B 區塊必須排在 app.js 之前，
      因為 app.js 尾端就會呼叫 refresh()，而 injectModule() 只查 DOM。 */
const firstTag = html.indexOf('<script src="');
const lastTag  = html.lastIndexOf('</script>') + '</script>'.length;
if (firstTag < 0) throw new Error('index.html 結構不符預期');

const appIdx = blocks.findIndex(b => b.startsWith('<!-- assets/js/app.js'));
const appPair = blocks.splice(appIdx, 2);                 /* 註解 + <script> 兩塊 */
blocks.push('');
blocks.push(...appPair);

html = html.slice(0, firstTag) + blocks.join('\n') + '\n' + html.slice(lastTag);

/* 4. 標題與說明改成單檔版口吻 */
html = html.replace(
  /<span class="badge on" id="badgeVer"><span class="dot"><\/span>[^<]*<\/span>/,
  '<span class="badge on" id="badgeVer"><span class="dot"></span>雙軌單檔版 ' +
  new Date().toISOString().slice(0,10) + '</span>');
html = html.replace(
  /<p><b>雙軌分流：<\/b>[\s\S]*?<\/p>/,
  '<p><b>雙軌分流：</b>本檔為單檔版，所有模組已內嵌。路徑 B 專屬的 '
  + '<code>price-calc</code>、<code>case-report</code>、<code>docx-writer</code>、'
  + '<code>docx-report</code>、<code>attachment-workbook</code>、<code>template</code> '
  + '以未執行的文字區塊存放，選擇路徑 B 時才執行，路徑 A 仍不會觸發試算邏輯。</p>');
html = html.replace(
  /<p><b>零外部依賴：<\/b>[\s\S]*?<\/p>/,
  '<p><b>零外部依賴：</b>自己解 <code>.xlsx</code>（<code>DecompressionStream</code> + XML 解析），'
  + '自己組 <code>.docx</code> / <code>.xlsx</code>（自寫 ZIP 打包器 + OOXML）。'
  + '單檔可直接以瀏覽器開啟，除字型 CDN 外不連外網，字型連不到時自動降階為系統字體。</p>');

fs.mkdirSync(OUT_DIR, { recursive: true });
fs.writeFileSync(OUT, html);

const kb = (fs.statSync(OUT).size / 1024).toFixed(0);
console.log('→ ' + path.relative(ROOT, OUT) + '  ' + kb + ' KB');
console.log('   共用模組 ' + staticSrcs.length + ' 支（內嵌執行）');
console.log('   路徑 B 模組 ' + pathB.length + ' 支（內嵌但延後執行）');
