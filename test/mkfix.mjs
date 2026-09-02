import fs from 'fs'; import vm from 'vm';
const R=(process.env.APP_ROOT || new URL('..',import.meta.url).pathname.replace(/\/$/,'')) + '/';
const ctx={console,Math,JSON,Date,TextEncoder,TextDecoder,Blob,Response,Promise,
  CompressionStream:globalThis.CompressionStream,DecompressionStream:globalThis.DecompressionStream,
  Uint8Array,Uint32Array,DataView,ArrayBuffer,Number,String,Object,Array,RegExp,Error,BigInt,Intl,
  parseInt,parseFloat,isNaN,isFinite,setTimeout,URL};
ctx.globalThis=ctx; ctx.window=ctx; vm.createContext(ctx);
for (const f of ['assets/js/format-utils.js','assets/js/parser.js','assets/js/xlsx-writer.js'])
  vm.runInContext(fs.readFileSync(R+f,'utf8'),ctx,{filename:f});
function row(o){ return Object.assign({
  CODE:'', 藥品名稱:'', 藥商名稱:'廠商甲', 分組代碼:'', 分組名稱:'', ATC7碼:'A02BA03', ATC名稱:'Famotidine',
  分類:'1', 藥品分類:'1', 藥品分類_名稱:'原開發廠藥品', 必要藥品:'3', 劑型:'錠劑', 適應症:'消化性潰瘍',
  收載日期:'19950301', 生效日期:'20180101', 最新申報年度資料範圍:'11501-11506',
  不良品暫停支付註記:'', 不上網註記:'', 藥價說明:'',
  PRICE11508:10, PRICE11509:10, QTY112:1000, QTY113:1100, QTY114:1200, AMT112:10000, AMT113:11000, AMT114:12000
}, o); }
const M=[
 row({CODE:'A001',藥品名稱:'甲錠20MG',分組代碼:'A02FAM211012',分組名稱:'FAMOTIDINE 20MG'}),
 row({CODE:'A002',藥品名稱:'乙錠20MG',分組代碼:'A02FAM211012',分組名稱:'FAMOTIDINE 20MG',PRICE11508:0,不良品暫停支付註記:'Y'}),
 row({CODE:'A003',藥品名稱:'丙錠20MG',分組代碼:'A02FAM211012',分組名稱:'FAMOTIDINE 20MG',PRICE11508:0,生效日期:'20120101'}),
 row({CODE:'A004',藥品名稱:'丁錠40MG',分組代碼:'A02FAM211014',分組名稱:'FAMOTIDINE 40MG',收載日期:'20240101',PRICE11508:18.5}),
 row({CODE:'TPN999',藥品名稱:'院內調製',分組代碼:'A02FAM211012',分組名稱:'FAMOTIDINE 20MG'}),
 row({CODE:'A005',藥品名稱:'不上網品',分組代碼:'A02FAM211012',分組名稱:'FAMOTIDINE 20MG',不上網註記:'Y'}),
 row({CODE:'B001',藥品名稱:'戊注射劑',分組代碼:'A02FAM211202',分組名稱:'FAMOTIDINE INJ',ATC7碼:'A02BA53',PRICE11508:120})
];
const C=[{CODE:'A001',ATC_MODE:'BOTH',報告人:'王小明',來文日期文號:'115年3月14日健保審字第1150001234號',
  藥品治療用途:'',廠商建議價:48.4,廠商成本:32.5,十國中位價:41,領有許可證:'Y',十國加計比例:'0,10',成本加計比例:'15,20,25,30'}];
const b1=await ctx.buildXlsx(M,'主檔'); fs.mkdirSync(R+'test/fixtures',{recursive:true}); fs.writeFileSync(R+'test/fixtures/fix_master.xlsx',Buffer.from(await b1.arrayBuffer()));
const b2=await ctx.buildXlsx(C,'案件'); fs.writeFileSync(R+'test/fixtures/fix_case.xlsx',Buffer.from(await b2.arrayBuffer()));
console.log('fixtures written', b1.size, b2.size);
