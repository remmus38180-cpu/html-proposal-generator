import fs from 'fs';
import vm from 'vm';
const R=(process.env.APP_ROOT || new URL('..',import.meta.url).pathname.replace(/\/$/,'')) + '/';
const ORDER=['assets/js/format-utils.js','assets/js/parser.js','assets/js/xlsx-writer.js','assets/js/group-model.js',
 'assets/js/exporters/sas-workbook.js','assets/js/exporters/query-workbook.js',
 'assets/js/price-calc.js','assets/js/case-report.js','assets/js/docx-writer.js',
 'assets/js/exporters/docx-report.js','assets/js/exporters/attachment-workbook.js','assets/js/template.js'];
const ctx={console,Math,JSON,Date,TextEncoder,TextDecoder,Blob,Response,Promise,
  CompressionStream:globalThis.CompressionStream,DecompressionStream:globalThis.DecompressionStream,
  Uint8Array,Uint32Array,DataView,ArrayBuffer,Number,String,Object,Array,RegExp,Error,BigInt,Intl,
  parseInt,parseFloat,isNaN,isFinite,setTimeout,URL};
ctx.globalThis=ctx; ctx.window=ctx;
vm.createContext(ctx);
for (const f of ORDER) vm.runInContext(fs.readFileSync(R+f,'utf8'), ctx, {filename:f});
console.log('modules loaded OK');

// ---- synthetic master ----
const M=[];
function row(o){ return Object.assign({
  CODE:'', 藥品名稱:'', 藥商名稱:'廠商甲', 分組代碼:'', 分組名稱:'', ATC7碼:'A02BA03', ATC名稱:'Famotidine',
  分類:'1', 藥品分類:'1', 藥品分類_名稱:'', 必要藥品:'3', 劑型:'錠劑', 適應症:'測試適應症',
  收載日期:'19950301', 生效日期:'20180101', 最新申報年度資料範圍:'11501-11506',
  PRICE11508:10, QTY112:1000, QTY113:1100, QTY114:1200, AMT112:10000, AMT113:11000, AMT114:12000
}, o); }
M.push(row({CODE:'A001',藥品名稱:'甲錠',分組代碼:'A02FAM211012',分組名稱:'FAMOTIDINE 20MG'}));
M.push(row({CODE:'A002',藥品名稱:'乙錠',分組代碼:'A02FAM211012',分組名稱:'FAMOTIDINE 20MG',PRICE11508:0,'不良品暫停支付註記':'Y'}));
M.push(row({CODE:'A003',藥品名稱:'丙錠',分組代碼:'A02FAM211012',分組名稱:'FAMOTIDINE 20MG',PRICE11508:0,生效日期:'20120101'}));
M.push(row({CODE:'A004',藥品名稱:'丁錠',分組代碼:'A02FAM211014',分組名稱:'FAMOTIDINE 40MG',PRICE11508:undefined,收載日期:'20240101'}));
M.push(row({CODE:'TPN999',藥品名稱:'排除品',分組代碼:'A02FAM211012',分組名稱:'FAMOTIDINE 20MG'}));
M.push(row({CODE:'A005',藥品名稱:'不上網品',分組代碼:'A02FAM211012',分組名稱:'FAMOTIDINE 20MG','不上網註記':'Y'}));

const ex=ctx.excludeMaster(M,true,true,{});
console.log('[A3] 排除 TPN',ex.tpn,'不上網',ex.offnet,'→ 採用',ex.rows.length);
const yr3=ctx.resolveYearWindow(ex.rows,115);
console.log('[A1] 最新申報年度資料範圍 11501-11506 (6個月) → YR3 =',yr3,'(期望 114)');
const model=ctx.buildGroupModel(ex.rows,'11508',yr3);
const g=model.groups['A02FAM211012'];
console.log('[B5] 分組 A02FAM211012 項目數 =',g.cnt,'/ 總數',g.total,'(A001價10✓ A002價0但Y✓ A003價0非Y✗ → 期望 2/3)');
console.log('[A2] A002(Y,價0,生效101) qtyAdj =',model.byCode['A002'].qtyAdj,'(期望不歸零)');
console.log('[A2] A003(非Y,價0,生效101) qtyAdj =',model.byCode['A003'].qtyAdj,'(期望 [0,0,0]，112-114 皆 >101)');
const g2=model.groups['A02FAM211014'];
console.log('[B6] 分組收載 113 年(20240101) → avgQty3 =',g2.avgQty3,'listYear',g2.listYear);
console.log('[B7] monthlyAmt =',g.monthlyAmt);
console.log('[B8] cutoffYear =',model.cutoffYear,'ingYears =',JSON.stringify(model.ingYears));
console.log('[B8] A001 cat3B =',model.byCode['A001'].cat3B,'ingListYear =',model.byCode['A001'].ingListYear);

// ---- 收載年晚於 YR3 ----
const M2=[row({CODE:'Z1',藥品名稱:'新藥',分組代碼:'N05ZZZ211020',分組名稱:'NEW',收載日期:'20260101',QTY112:0,QTY113:0,QTY114:0,AMT112:0,AMT113:0,AMT114:0})];
const mz=ctx.buildGroupModel(M2,'11508',114);
console.log('[B6] 收載年 115 > YR3 114 → avgQty3 =',mz.groups['N05ZZZ211020'].avgQty3,'(期望 0，不是 null)');

// ---- [D19] ----
for (const v of [823,50000,123456789,1234,10020000000,1500000000000])
  console.log('[D19] chineseAmount2('+v+') =',ctx.chineseAmount2(v),'| chineseAmount =',ctx.chineseAmount(v,0));
// ---- [C11]/[C12] ----
console.log('[C12] truncatePrice 3.8*1.2 =',ctx.truncatePrice(3.8*1.2),'(期望 4.56)');
console.log('[C12] truncatePrice(4.55999) =',ctx.truncatePrice(4.55999),'(期望 4.55)');
console.log('[C10] costRefPrice(32.5,0.5,"Y") =',ctx.costRefPrice(32.5,0.5,'Y'));
console.log('[C10] costRefPrice(32.5,0.5,"N") =',ctx.costRefPrice(32.5,0.5,'N'));
console.log('[D17] essentialLabel(3) =',ctx.essentialLabel('3'),'| (2) =',ctx.essentialLabel('2'));
console.log('[D15] drugClassLabel({藥品分類:2}) =',ctx.drugClassLabel({'藥品分類':'2'}));

// ---- Path A workbook ----
const q=ctx.buildQueryWorkbook(model,{atc7:[],grpCd:[],grpNm:[],code:[],brand:[]},'11508');
console.log('PathA counts',JSON.stringify(q.counts));
const qb=await q.book; console.log('PathA xlsx bytes',qb.size);

// ---- Path B ----
const cases=[{CODE:'A001',ATC_MODE:'BOTH',報告人:'王小明',來文日期文號:'115年3月14日健保審字第1150001234號',
  藥品治療用途:'',廠商建議價:48.4,廠商成本:32.5,十國中位價:41,領有許可證:'Y',十國加計比例:'0,10',成本加計比例:'15,20,25,30'}];
const rep=ctx.buildCaseReport(cases,model,'11508');
console.log('PathB rep items',rep.items.length,'table1',rep.table1.length,'blocks',rep.table2Blocks.length);
console.log('PathB calc priceRow',rep.items[0].calc.priceRow);
console.log('PathB calc totalRow',rep.items[0].calc.totalRow);
const sb=await ctx.buildSasWorkbook(model,{atc7:[],grpCd:[],grpNm:[],code:[],brand:[]},rep.items,'11508',{skipPriceCalc:false}).book;
console.log('PathB 四分頁 bytes',sb.size);
const sbA=ctx.buildSasWorkbook(model,{atc7:[],grpCd:[],grpNm:[],code:[],brand:[]},[],'11508',{skipPriceCalc:true});
console.log('skipPriceCalc=true →',(await sbA.book).size,'bytes, skip flag',sbA.skipPriceCalc);
const dx=await ctx.reportToDocx(rep); console.log('PathB docx bytes',dx.size);
const at=await ctx.buildAttachmentWorkbook(model,rep,'11508',[4,5,7]).book; console.log('PathB 附件 bytes',at.size);
console.log('ALL OK');
