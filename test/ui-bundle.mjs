import { chromium } from 'playwright';
import fs from 'fs'; import os from 'os'; import path from 'path';
const ROOT=process.env.APP_ROOT || path.resolve(new URL('..',import.meta.url).pathname);
/* 刻意複製到一個「只有這一個檔」的空目錄，證明單檔真的不依賴 assets/ */
const DIR=fs.mkdtempSync(path.join(os.tmpdir(),'onefile-'));
fs.copyFileSync(ROOT+'/dist/case_template.html', DIR+'/case_template.html');
const FIX=ROOT+'/test/fixtures';
const b=await chromium.launch(process.env.CHROME ? {executablePath:process.env.CHROME} : {});
const pg=await b.newPage({acceptDownloads:true});
const errs=[], reqs=[];
pg.on('console',m=>{ if(m.type()==='error') errs.push(m.text()); });
pg.on('pageerror',e=>errs.push('PAGEERROR: '+e.message));
pg.on('request',r=>{ if(/\.js(\?|$)/.test(r.url())) reqs.push(r.url()); });
// 用 file:// 開，且目錄裡「只有」這一個檔
await pg.goto('file://'+DIR+'/case_template.html');
await pg.waitForTimeout(400);
console.log('file:// 開檔，額外 .js 請求:', reqs.length ? reqs.join(' ') : '(無)');
console.log('引擎徽章:', await pg.textContent('#badgeEngine'));

await pg.setInputFiles('#fMaster', FIX+'/fix_master.xlsx');
await pg.waitForTimeout(700);
console.log('主檔:', await pg.textContent('#sMaster'));
console.log('PRICE 下拉:', await pg.$$eval('#pySelect option',o=>o.map(x=>x.textContent).join(' | ')));
await pg.selectOption('#pySelect','11508'); await pg.waitForTimeout(300);
console.log('支付價欄位:', await pg.textContent('#dPriceCol'), '| 視窗:', await pg.textContent('#dDataYear'), '| B8基準年:', await pg.textContent('#dCutoff'));

// 路徑 A：確認 Path B 的函式在此時「不存在」
await pg.check('#rdA'); await pg.waitForTimeout(400);
console.log('選 A 後 typeof buildCaseReport =', await pg.evaluate(()=>typeof buildCaseReport));
console.log('選 A 後 typeof costRefPrice   =', await pg.evaluate(()=>typeof costRefPrice));
console.log('選 A 後 typeof reportToDocx   =', await pg.evaluate(()=>typeof reportToDocx));
let d=pg.waitForEvent('download'); await pg.click('#btnQuery');
await (await d).saveAs(DIR+'/one_query.xlsx'); console.log('PathA 查詢表已下載');

// 路徑 B：延後執行的區塊此時才跑
await pg.check('#rdB'); await pg.waitForTimeout(900);
console.log('選 B 後 typeof buildCaseReport =', await pg.evaluate(()=>typeof buildCaseReport));
console.log('選 B 後 typeof costRefPrice   =', await pg.evaluate(()=>typeof costRefPrice));
console.log('載入路徑B時的 .js 網路請求:', reqs.length ? reqs.join(' ') : '(無，全部內嵌)');
await pg.setInputFiles('#fCase', FIX+'/fix_case.xlsx'); await pg.waitForTimeout(600);
console.log('案件檔:', await pg.textContent('#sCase'));
await pg.click('#btnPreview'); await pg.waitForTimeout(400);
console.log('預覽表格數:', await pg.$$eval('#prevBox table',t=>t.length));
for (const [btn,lab,ext] of [['#btnDocx','Word','.docx'],['#btnXlsx','四分頁','.xlsx'],['#btnAttach','附件','.xlsx'],['#btnTpl','範本','.xlsx']]){
  const dp=pg.waitForEvent('download'); await pg.click(btn);
  await (await dp).saveAs(DIR+'/one_'+lab+ext); console.log('PathB 下載:', lab);
  await pg.waitForTimeout(300);
}
console.log('--- JS 錯誤:', errs.length ? errs.join('\n') : '(無)');
await b.close();
fs.rmSync(DIR,{recursive:true,force:true});
