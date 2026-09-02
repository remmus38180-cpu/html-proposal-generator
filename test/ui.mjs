import { chromium } from 'playwright';
import http from 'http'; import fs from 'fs'; import path from 'path';
const ROOT=process.env.APP_ROOT || path.resolve(new URL('..',import.meta.url).pathname);
const FIX=process.env.FIX_DIR || ROOT+'/test/fixtures';
const MT={'.html':'text/html','.js':'text/javascript','.css':'text/css'};
const srv=http.createServer((q,s)=>{
  let p=path.join(ROOT, decodeURIComponent(q.url.split('?')[0]));
  if (p.endsWith('/')) p+='index.html';
  fs.readFile(p,(e,d)=>{ if(e){s.writeCode=404;s.writeHead(404);s.end('404');} else {s.writeHead(200,{'Content-Type':MT[path.extname(p)]||'application/octet-stream'});s.end(d);} });
});
await new Promise(r=>srv.listen(8731,r));
const b=await chromium.launch(process.env.CHROME ? {executablePath:process.env.CHROME} : {});
const pg=await b.newPage();
const errs=[], reqs=[];
pg.on('console',m=>{ if(m.type()==='error') errs.push(m.text()); });
pg.on('pageerror',e=>errs.push('PAGEERROR: '+e.message));
pg.on('request',r=>{ if(r.url().endsWith('.js')) reqs.push(r.url().replace('http://localhost:8731/','')); });
await pg.goto('http://localhost:8731/index.html');
await pg.waitForTimeout(300);
console.log('diagBox 初始可見:', await pg.isVisible('#diagBox'));
console.log('初次載入的 JS:', reqs.join(' '));

// 1. 上傳主檔
await pg.setInputFiles('#fMaster',FIX+'/fix_master.xlsx');
await pg.waitForTimeout(600);
console.log('主檔狀態:', await pg.textContent('#sMaster'));
console.log('PRICE 下拉:', await pg.$$eval('#pySelect option',o=>o.map(x=>x.textContent).join(' | ')));
console.log('路徑步驟可見(選年月前):', await pg.isVisible('#stepTrack'));

// 2. 錯誤年月
await pg.fill('#priceYear','11501'); await pg.waitForTimeout(200);
console.log('11501 → 支付價欄位:', await pg.textContent('#dPriceCol'), '| 路徑步驟:', await pg.isVisible('#stepTrack'));

// 3. 正確年月
await pg.selectOption('#pySelect','11508'); await pg.waitForTimeout(300);
console.log('11508 → 支付價欄位:', await pg.textContent('#dPriceCol'));
console.log('近三年視窗:', await pg.textContent('#dDataYear'), '| [B8]基準年:', await pg.textContent('#dCutoff'));
console.log('已排除:', await pg.textContent('#dExcl'));
console.log('路徑步驟可見:', await pg.isVisible('#stepTrack'));

// 4. 路徑 A
reqs.length=0;
await pg.check('#rdA'); await pg.waitForTimeout(400);
console.log('選 A 後新載入的 JS:', reqs.length? reqs.join(' ') : '(無)');
console.log('stepOutA/stepOutB 可見:', await pg.isVisible('#stepOutA'), await pg.isVisible('#stepOutB'));
const dl1=pg.waitForEvent('download');
await pg.click('#btnQuery');
const d1=await dl1; await d1.saveAs(FIX+'/out_query.xlsx'); console.log('PathA 下載:', d1.suggestedFilename(), '→ out_query.xlsx');
await pg.waitForTimeout(200);

// 5. 切到 B（有 confirm 但無 caseData → 不應跳）
reqs.length=0;
await pg.check('#rdB'); await pg.waitForTimeout(900);
console.log('選 B 後新載入的 JS:', reqs.join(' '));
console.log('casePanel 可見:', await pg.isVisible('#casePanel'));
await pg.setInputFiles('#fCase',FIX+'/fix_case.xlsx');
await pg.waitForTimeout(600);
console.log('案件檔狀態:', await pg.textContent('#sCase'), '| 本案藥品:', await pg.textContent('#dCaseN'));
await pg.click('#btnPreview'); await pg.waitForTimeout(400);
console.log('預覽表格數:', await pg.$$eval('#prevBox table',t=>t.length));
for (const [btn,label] of [['#btnDocx','Word'],['#btnXlsx','四分頁'],['#btnAttach','附件']]){
  const dp=pg.waitForEvent('download'); await pg.click(btn);
  const dd=await dp; await dd.saveAs(FIX+'/out_'+label+(label==='Word'?'.docx':'.xlsx')); console.log('PathB 下載('+label+'):', dd.suggestedFilename());
  await pg.waitForTimeout(300);
}
// 6. B → A 需確認
pg.on('dialog',async d=>{ console.log('confirm 對話框:', d.message()); await d.accept(); });
await pg.check('#rdA'); await pg.waitForTimeout(400);
console.log('切回 A 後 casePanel 可見:', await pg.isVisible('#casePanel'), '| sCase:', await pg.textContent('#sCase'));

console.log('--- console log 內容 ---');
console.log(await pg.textContent('#logBox'));
console.log('--- JS 錯誤:', errs.length ? errs.join('\n') : '(無)');
await b.close(); srv.close();
