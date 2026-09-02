'use strict';
/* 分組模型：[A2]/[A3]/[B5]~[B8] 共用核心（Path A / Path B 共用） */

/* --- ATC_資訊 --- */
function atcInfo(d){
  var a = atcOf(d), n = txt(d['ATC名稱']).toLowerCase();
  return a ? ('[' + a + '] ' + n) : n;
}

/* --- 不良品暫停支付註記（當月欄位無年月後綴） --- */
function suspFlag(d){
  return String(d['不良品暫停支付註記'] || '').trim().toUpperCase() === 'Y';
}

function ingredientKey(v){
  var s = txt(v);
  if (s.length < 11) return '';
  return s.substring(0,8) + s.substring(9,11);
}

/* --- 支付價區間（含0元，SAS：distinct 排序、' / ' 串接、非整數去尾零） --- */
function priceRange(items, priceCol, sep, skipZero){
  var seen = {}, out = [];
  for (var i=0;i<items.length;i++){
    var v = num(items[i][priceCol]);
    if (v === null) continue;
    if (skipZero && v === 0) continue;
    if (seen[v]) continue;
    seen[v] = 1; out.push(v);
  }
  out.sort(function(a,b){ return a-b; });
  return out.map(function(v){ return String(v); }).join(sep || ' / ');
}

function detectDataYear(master, baseYear){
  var keys = {}, i, k;
  for (i=0;i<Math.min(master.length,300);i++) for (k in master[i]) keys[k]=1;
  for (var y = baseYear-1; y >= baseYear-8; y--) if (keys['QTY'+y] && keys['AMT'+y]) return y;
  return baseYear-1;
}

function qtyOf(d,y){ return num(d['QTY'+y])||0; }

function amtOf(d,y){ return num(d['AMT'+y])||0; }

function atcOf(d){ return String(d['ATC7碼']===undefined||d['ATC7碼']===null?'':d['ATC7碼']).trim(); }

function atcMatch(d, atc7, mode){
  var a = atcOf(d);
  return mode === 'PREFIX5' ? a.substring(0,5) === atc7.substring(0,5) : a === atc7;
}

/* [B5] 同分組項目數（給付中，不含已取消支付價項目）
   1. 不良品暫停支付註記 = 'Y'          → 計入（無論支付價為何）
   2. 註記 ≠ 'Y' 且 支付價 ≠ 0          → 計入
   3. 註記 ≠ 'Y' 且 支付價為空值        → 計入（費用內含於其他項目，不另核價但有申報量）
   唯一不計入的情形：註記 ≠ 'Y' 且 支付價 = 0 */
function countValid(items, priceCol){
  var n = 0;
  for (var i=0;i<items.length;i++){
    if (num(items[i][priceCol]) === 0 && !suspFlag(items[i])) continue;
    n++;
  }
  return n;
}

/* 價格 0 但不良品暫停支付註記 = Y → 計入項目數，但支付價欄不顯示 */
function countSuspended(items, priceCol){
  var n = 0;
  for (var i=0;i<items.length;i++){
    if (num(items[i][priceCol]) === 0 && suspFlag(items[i])) n++;
  }
  return n;
}

/* 【列出／計數的唯一判定】與 countValid 同一條 [B5] 規則，只是以「整備後的列」為輸入。
   這條規則只決定「哪些列要列出／計數」；金額與數量的採計一律走 [A2] 的
   amtAdj／qtyAdj，兩者互不影響。 */
function isCounted(r){
  if (r.price === 0) return suspFlag(r.d);
  return true;                       /* 支付價 ≠ 0 或為空值 → 均計入 */
}

function priceList(items, priceCol){
  var seen = {}, out = [];
  for (var i=0;i<items.length;i++){
    var v = num(items[i][priceCol]);
    if (v === null || v === 0) continue;
    if (seen[v]) continue;
    seen[v] = 1; out.push(v);
  }
  out.sort(function(a,b){return a-b;});
  return out.map(money).join('、');
}

function listedYears(drug, baseYear){
  var raw = drug['收載日期'];
  if (raw === undefined || raw === null || raw === '') return null;
  if (typeof raw === 'number' && raw > 20000 && raw < 80000){          // Excel 日期序號
    var dt = new Date(Date.UTC(1899,11,30) + raw*86400000);
    return baseYear - (dt.getUTCFullYear() - 1911);
  }
  var s = String(raw).replace(/\D/g,''), y = null;
  if (s.length === 7) y = parseInt(s.substring(0,3),10);
  else if (s.length === 8) y = parseInt(s.substring(0,4),10) - 1911;
  if (y === null || isNaN(y)) return null;
  return baseYear - y;
}

/* [A1] 解析主檔的 {最新申報年度資料範圍}（例 '11501-11506' / '11501-11512'）
   回傳 {year:民國年, months:涵蓋月數}；無法解析回 null。 */
var DATARANGE_KEYS = ['最新申報年度資料範圍','最新申報年度資料範圍 ','申報年度資料範圍'];
function parseDataRange(master){
  var i, k, raw = '';
  for (i=0;i<Math.min(master.length,300) && !raw;i++)
    for (k=0;k<DATARANGE_KEYS.length && !raw;k++) raw = txt(master[i][DATARANGE_KEYS[k]]);
  if (!raw) return null;
  var m = String(raw).match(/(\d{3})(\d{2})\s*[-~—－至]\s*(\d{3})(\d{2})/);
  if (!m) return null;
  var y1 = parseInt(m[1],10), m1 = parseInt(m[2],10),
      y2 = parseInt(m[3],10), m2 = parseInt(m[4],10);
  if (isNaN(y1)||isNaN(m1)||isNaN(y2)||isNaN(m2)) return null;
  var months = (y2 - y1) * 12 + (m2 - m1) + 1;
  return {year:y2, months:months, raw:String(raw)};
}

/* [A1] 年度視窗：當期年 YR3 = 有完整12個月申報資料的最近一年
   1. 主檔有 {最新申報年度資料範圍} → 涵蓋 12 個月取該民國年，不足 12 個月往前推 1 年
   2. 無此欄位 → 退回舊行為：自 baseYear-1 起往前找存在 QTY/AMT 欄的年度（最多回推 8 年）
   兩種來源最後都會再確認該年度的 QTY/AMT 欄位確實存在。 */
function resolveYearWindow(master, baseYear){
  var keys = {}, i, k;
  for (i=0;i<Math.min(master.length,300);i++) for (k in master[i]) keys[k]=1;
  var has = function(y){ return !!(keys['QTY'+y] && keys['AMT'+y]); };

  var dr = parseDataRange(master);
  if (dr){
    var y = (dr.months >= 12) ? dr.year : dr.year - 1;
    if (has(y)) return y;
    for (; y >= baseYear-8; y--) if (has(y)) return y;   /* 欄位真的缺就往前退 */
  }
  for (var z = baseYear-1; z >= baseYear-8; z--) if (has(z)) return z;
  return baseYear-1;
}

var OFFNET_KEYS = ['不上網註記','不上網','不上網註','不上網記號'];

function offnetFlag(d){
  for (var i=0;i<OFFNET_KEYS.length;i++){
    var v = txt(d[OFFNET_KEYS[i]]);
    if (v) return v.toUpperCase();
  }
  return '';
}

function isTpnCode(code){ return txt(code).toUpperCase().substring(0,3) === 'TPN'; }

function excludeMaster(rows, exTpn, exOff, keepCodes){
  var out = [], nT = 0, nO = 0;
  for (var i=0;i<rows.length;i++){
    var d = rows[i], code = txt(d['CODE']);
    if (!(keepCodes && keepCodes[code])){
      if (exTpn && isTpnCode(code)){ nT++; continue; }
      if (exOff && offnetFlag(d) === 'Y'){ nO++; continue; }
    }
    out.push(d);
  }
  return {rows: out, tpn: nT, offnet: nO};
}

function buildGroupModel(master, priceYear, yr3){
  var baseYear = parseInt(priceYear.substring(0,3),10);
  var priceCol = 'PRICE' + priceYear;
  var YR3 = (yr3 === undefined || yr3 === null) ? baseYear-1 : yr3;
  var YR1 = YR3-2, YR2 = YR3-1;
  var i, k;

  /* [A2] 逐列：0元項目的年度歸零
     情形1 不良品暫停支付註記 = 'Y'          → 直接採計，無年份邊界（無論支付價為何）
     情形2 註記 ≠ 'Y' 且 支付價 ≠ 0          → 直接採計
     情形3 註記 ≠ 'Y' 且 支付價 = 0          → 年份 ≤ 生效年 採計，大於者歸零
     情形4 註記 ≠ 'Y' 且 支付價為空值        → 直接採計（費用內含於其他項目，無價仍有申報量） */
  var rows = master.map(function(d){
    var p = num(d[priceCol]);
    var eff = ad8(d['生效日期']);
    var susp = suspFlag(d);
    var q = [num(d['QTY'+YR1])||0, num(d['QTY'+YR2])||0, num(d['QTY'+YR3])||0];
    var a = [num(d['AMT'+YR1])||0, num(d['AMT'+YR2])||0, num(d['AMT'+YR3])||0];
    var qa = q.slice(), aa = a.slice();
    if (!susp && p === 0 && eff !== null){
      var yrs = [YR1,YR2,YR3];
      for (var j=0;j<3;j++) if (!(yrs[j] <= eff)){ qa[j] = 0; aa[j] = 0; }
    }
    return {d:d, price:p, susp:susp, code:txt(d['CODE']), grp:txt(d['分組代碼']),
            ingKey:ingredientKey(d['分組代碼']),
            qty:q, amt:a, qtyAdj:qa, amtAdj:aa};
  });

  /* [B8] 同成分同劑型收載年（10碼 key，與 [B6] 的12碼分組各自獨立 GROUP BY）
     基準年（藥商銷售資料採計期間截止年）：price 年月之月份 ≥ 4 → 該民國年；< 4 → 該民國年 − 1
     基準年 − 同成分同劑型收載年 ≥ 15 → 第三B大類 */
  var mm = parseInt(priceYear.substring(3,5),10);
  var cutoffYear = (isNaN(mm) || mm >= 4) ? baseYear : baseYear - 1;
  var ING = {};
  for (i=0;i<rows.length;i++){
    var ik = rows[i].ingKey;
    if (!ik) continue;
    var iy = ad8(rows[i].d['收載日期']);
    if (iy === null) continue;
    if (ING[ik] === undefined || iy < ING[ik]) ING[ik] = iy;
  }
  for (i=0;i<rows.length;i++){
    var y8 = (rows[i].ingKey && ING[rows[i].ingKey] !== undefined) ? ING[rows[i].ingKey] : null;
    rows[i].ingListYear = y8;
    rows[i].cat3B = (y8 === null) ? null : (cutoffYear - y8 >= 15);
  }

  /* 依分組代碼彙總 */
  var G = {}, gOrder = [];
  for (i=0;i<rows.length;i++){
    var g = rows[i].grp;
    if (!G[g]){
      G[g] = {code:g, name:txt(rows[i].d['分組名稱']), items:[], rows:[],
              qtySum:[0,0,0], amtSum:[0,0,0], listYear:null};
      gOrder.push(g);
    }
    var o = G[g];
    o.rows.push(rows[i]); o.items.push(rows[i].d);
    for (k=0;k<3;k++){ o.qtySum[k] += rows[i].qtyAdj[k]; o.amtSum[k] += rows[i].amtAdj[k]; }
    var ly = ad8(rows[i].d['收載日期']);
    if (ly !== null && (o.listYear === null || ly < o.listYear)) o.listYear = ly;
  }
  /* [B6] 同分組近三年平均申報量
     分母由分組收載年決定，固定為 3 / 2 / 1，不隨 NULL 年度縮減：
       收載年 早於 YR1（近第3年）      → ÷3
       收載年 = YR1 / YR2 / YR3        → ÷3 / ÷2 / ÷1
       收載年 晚於 YR3（近第1年）      → 該分組近三年均無申報資料，平均為 0（仍須列出）
     [B7] 同分組每月申報金額 = 當期年度（YR3）AMT 總計 ÷ 12 */
  var yrs = [YR1, YR2, YR3];
  for (i=0;i<gOrder.length;i++){
    var gg = G[gOrder[i]], vals = [];
    for (k=0;k<3;k++){
      if (gg.listYear !== null && yrs[k] < gg.listYear) continue;
      vals.push(Math.round(fpClean(gg.qtySum[k])*10)/10);
    }
    gg.avgQty3 = vals.length
      ? Math.round(fpClean(vals.reduce(function(a,b){return a+b;},0)/vals.length)*10)/10
      : 0;                       /* 收載年晚於近第1年 → 0，不是 null */
    gg.amtSum  = gg.amtSum.map(function(v){ return Math.round(fpClean(v)*10)/10; });
    gg.qtySum  = gg.qtySum.map(function(v){ return Math.round(fpClean(v)*10)/10; });
    gg.monthlyAmt = Math.round(fpClean(gg.amtSum[2]/12)*10)/10;
    gg.priceRange = priceRange(gg.items, priceCol, ' / ');
    gg.cnt = countValid(gg.items, priceCol);       /* [B5] 給付中項目數 */
    gg.total = gg.rows.length;
    /* [B8] 同一分組代碼之所有項目共用同一個10碼 key，取第一列即可 */
    gg.ingKey      = gg.rows[0].ingKey;
    gg.ingListYear = gg.rows[0].ingListYear;
    gg.cat3B       = gg.rows[0].cat3B;
  }
  var byCode = {};
  for (i=0;i<rows.length;i++) byCode[rows[i].code] = rows[i];

  return {baseYear:baseYear, priceCol:priceCol, YR1:YR1, YR2:YR2, YR3:YR3,
          cutoffYear:cutoffYear, ingYears:ING,
          rows:rows, groups:G, gOrder:gOrder, byCode:byCode, yrs:yrs};
}
