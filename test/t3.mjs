import fs from 'fs'; import vm from 'vm';
const ctx={console,Math,Number,String,Object,Array,JSON,BigInt,Intl,parseInt,parseFloat,isNaN,isFinite,Date,RegExp};
ctx.globalThis=ctx; vm.createContext(ctx);
vm.runInContext(fs.readFileSync((process.env.APP_ROOT || new URL('..',import.meta.url).pathname.replace(/\/$/,'')) + '/assets/js/format-utils.js','utf8'),ctx);
const T=[[3.8*1.2,4.56],[4.55999,4.55],[4.999999999,4.99],[5,5],[49.99,49.9],[50,50],[50.9,50],
 [0.001,0],[0.126,0.12],[12.345,12.3],[123.9,123],[1.005*100/100,1]];
for (const [v,e] of T){const r=ctx.truncatePrice(v);console.log('truncatePrice('+v+') =',r,'期望',e,r===e?'✓':'✗');}
console.log('--- truncate2 ---');
for (const [v,e] of [[32.5*1.5,48.75],[4.559999999999999,4.55],[1.0/3,0.33],[48.75*1.0505,51.21]]){
  const r=ctx.truncate2(v);console.log('truncate2('+v+') =',r,'期望',e,r===e?'✓':'✗');}
console.log('--- 邊界/防呆 ---');
for (const v of [null,undefined,NaN,'',0,-3.789,Infinity])
  console.log('truncatePrice('+String(v)+') =',ctx.truncatePrice(v));
console.log('chineseAmount2 邊界:',[0,-1234,9999,10000,9999999,10000000,99999999,100000000,999999999999,1000000000000].map(v=>v+'→'+ctx.chineseAmount2(v)).join('  '));
console.log('chineseAmount 防呆:',[null,'',NaN,-0.4].map(v=>JSON.stringify(ctx.chineseAmount(v,0))).join(' '));
console.log('formatAmt:',[0,823,50000,123456789].map(v=>v+'→'+ctx.formatAmt(v)).join('  '));
