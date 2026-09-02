'use strict';
/* 格式與數值工具（Path A / Path B 共用） */

function esc(s){
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function fpClean(v){
  if (typeof v !== 'number' || !isFinite(v)) return v;
  return Number(v.toPrecision(15));
}

/* 無條件捨去到指定小數位（已消除浮點誤差） */
function floorAt(v, dec){
  var f = Math.pow(10, dec);
  return Math.floor(fpClean(v * f)) / f;
}

/* ------------------------------------------------------------------
   [C11] 浮點誤差防護（處理原則版：toPrecision(10)）
   僅用於「價格」的截斷路徑。分組層金額／數量總計不走這裡 —— 兆級總計
   的有效位數超過 10 位，用 toPrecision(10) 會把真實數值一起磨掉，
   那條路徑維持 fpClean()（toPrecision(15)）。
   ------------------------------------------------------------------ */
/* ⚠️ 乘上級距倍數後會再生一次尾差（4.56 × 100 = 455.99999999999994），
   因此縮放前後都要清一次，否則 3.8×1.2 會被誤截為 4.55（正解 4.56）。
   處理原則所附的參考碼只清縮放前那一次。 */
function clean10(v){ return parseFloat(Number(v).toPrecision(10)); }
function floorScaled(v, scale){ return Math.floor(clean10(clean10(v) * scale)) / scale; }
/* 空字串、null、NaN、±Infinity 一律視為「無值」，不可讓 Number('')→0 混進核價 */
function priceNum(p){
  if (p === null || p === undefined || p === '') return null;
  var v = Number(p);
  return Number.isFinite(v) ? v : null;
}

function truncatePrice(p){
  var v = priceNum(p);
  if (v === null) return null;
  v = clean10(v);                                   // 消除二進位尾差
  if (v < 5)  return floorScaled(v, 100);
  if (v < 50) return floorScaled(v, 10);
  return Math.floor(v);
}
/* 成本法步驟2：固定取小數後兩位，第三位（含）以後無條件捨去 */
function truncate2(p){
  var v = priceNum(p);
  return v === null ? null : floorScaled(v, 100);
}

/* ------------------------------------------------------------------
   [D19] 金額顯示級距（簡化顯示）
   x<1萬 實數加千分位／1萬~<1千萬 萬(1位)／1千萬~<1億 億(2位)／
   1億~<1千億 億(1位)／1千億~<1兆 兆(2位)／≥1兆 兆(1位)
   ------------------------------------------------------------------ */
function chineseAmount2(val){
  if (val === null || val === undefined || val === '' || Number.isNaN(Number(val))) return '';

  const raw = Number(val);
  if (!Number.isFinite(raw)) return '數值超出範圍';
  const n = Math.round(raw);
  if (!Number.isSafeInteger(n)) {
    console.warn('chineseAmount2: 數值 ' + val + ' 超過安全整數範圍 (±' + Number.MAX_SAFE_INTEGER + ')，結果可能不精確');
  }
  if (n === 0) return '0元';

  const neg = n < 0;
  const av = BigInt(Math.abs(n));

  const WAN  = 10000n, QWAN = 10000000n, YI = 100000000n,
        QYI  = 100000000000n, ZHAO = 1000000000000n;

  const qf = (v, divisor, decimals) => {
    const scale = 10n ** BigInt(decimals);
    let q = v / divisor;
    const r = v % divisor;
    const scaled = r * scale;
    let f = scaled / divisor;
    const rem2 = scaled % divisor;
    if (rem2 * 2n >= divisor) f += 1n;
    if (f === scale) { q += 1n; f = 0n; }
    return { q, f };
  };
  const fmtq = (q, f, decimals, unit) =>
    q.toLocaleString('en-US') + '.' + f.toString().padStart(decimals, '0') + unit;

  let out;
  if (av < WAN) {
    out = av.toLocaleString('en-US') + '元';
  } else if (av < QWAN) {
    const { q, f } = qf(av, WAN, 1);
    out = fmtq(q, f, 1, '萬');
  } else if (av < YI) {
    const { q, f } = qf(av, YI, 2);
    out = fmtq(q, f, 2, '億');
  } else if (av < QYI) {
    const { q, f } = qf(av, YI, 1);
    out = fmtq(q, f, 1, '億');
  } else if (av < ZHAO) {
    const { q, f } = qf(av, ZHAO, 2);
    out = fmtq(q, f, 2, '兆');
  } else {
    const { q, f } = qf(av, ZHAO, 1);
    out = fmtq(q, f, 1, '兆');
  }
  return neg ? '-' + out : out;
}

/* ------------------------------------------------------------------
   [D19] 完整中文金額展開式（適用正式文件或需精確金額之場合）
   decimals 預設 3（範圍 0~6）；unit 預設「元」
   ------------------------------------------------------------------ */
function chineseAmount(val, decimals, unit){
  if (val === null || val === undefined || val === '') return '';

  const raw = Number(val);

  if (Number.isNaN(raw)) return '';
  if (!Number.isFinite(raw)) return '數值超出範圍';

  if (decimals === null || decimals === undefined || decimals === '') decimals = 3;
  decimals = Math.floor(Number(decimals));
  if (!Number.isFinite(decimals)) decimals = 3;
  if (decimals < 0) decimals = 0;
  if (decimals > 6) decimals = 6;

  if (unit === null || unit === undefined || unit === '') unit = '元';

  if (raw === 0) return '0' + unit;

  const neg = raw < 0;
  const absv = Math.abs(raw);

  let intv;
  let decstr = '';

  if (decimals === 0) {
    intv = Math.round(absv);
  } else {
    intv = Math.floor(absv);
    const scale = 10 ** decimals;
    let decvInt = Math.round((absv - intv) * scale);
    if (decvInt >= scale) { intv += 1; decvInt = 0; }
    if (decvInt > 0) decstr = '.' + String(decvInt).padStart(decimals, '0');
  }

  if (intv === 0 && decstr === '') return '0' + unit;   /* 避免 -0 */

  if (!Number.isSafeInteger(intv)) {
    console.warn('chineseAmount: 數值 ' + val + ' 的整數部分超過 JavaScript 安全整數範圍，結果可能不精確');
  }

  let rest = BigInt(intv);

  const zhao = rest / 1000000000000n; rest %= 1000000000000n;
  const yi   = rest / 100000000n;     rest %= 100000000n;
  const wan  = rest / 10000n;         rest %= 10000n;
  const yu   = rest;

  const f = x => x.toLocaleString('en-US');

  let out = '';
  if (zhao > 0n) out += f(zhao) + '兆';
  if (yi   > 0n) out += f(yi)   + '億';
  if (wan  > 0n) out += f(wan)  + '萬';
  if (yu > 0n || decstr !== '' || out === '') out += f(yu) + decstr;
  out += unit;

  return neg ? '-' + out : out;
}

/* 金額顯示級距（[D19] 簡化式），不帶單位「元」的變體供表格使用 */
function formatAmt(v){
  var s = chineseAmount2(v);
  return s.slice(-1) === '元' ? s.slice(0, -1) : s;
}

/* --- 藥品分類_名稱：代碼 → 提案用語 --- */
var DRUGCLASS_LABEL = {
  '1': '原開發廠藥品',
  '2': '生物製劑',
  '3': 'BA/BE學名藥品',
  '4': '一般學名藥品',
  '5': '生物相似性藥品',
  '9': 'BE對照品'
};

function drugClassLabel(d){
  var k = String(txt(d['藥品分類'])).trim();
  return DRUGCLASS_LABEL[k] || txt(d['藥品分類_名稱']) || '';
}

/* --- [D17] {必要藥品} 代碼對照 --- */
var ESSENTIAL_LABEL = {
  '0': '一般藥品',
  '1': '一般藥品',
  '3': '不可替代特殊藥品',
  '4': '特殊藥品'
};

function essentialLabel(v){
  var k = txt(v);
  if (k === '') return '';
  return ESSENTIAL_LABEL[k] || k;
}

/* --- 分類 → 第◯大類 --- */
var CAT_ORD = {'1':'第一','2':'第二','3':'第三A','4':'第三A','5':'第三B','6':'第三B'};

function catOrdinal(v){
  var k = String(v === undefined || v === null ? '' : v).trim();
  return CAT_ORD[k] || '';
}

/* --- 自適應百分比 --- */
function adaptivePct(v){
  if (v === null || v === undefined || isNaN(v)) return '';
  if (v === 0) return '0%';
  var a = Math.abs(v), dec;
  if (a >= 10) dec = 0;
  else if (a >= 1) dec = 1;
  else { dec = 0; var t = a; while (t < 1 && dec < 20){ t *= 10; dec++; } }
  return v.toFixed(dec) + '%';
}

/* --- 西元8碼 → 民國 YYY-MM-DD --- */
function toRocDate(v){
  var y = ad8(v);
  if (y === null) return '';
  var s = String(Math.round(rawAd8(v)));
  while (s.length < 8) s = '0' + s;
  return (y) + '-' + s.substring(4,6) + '-' + s.substring(6,8);
}

function rawAd8(v){
  if (v === undefined || v === null || v === '') return null;
  if (typeof v === 'number' && v > 20000 && v < 80000){            // Excel 日期序號
    var dt = new Date(Date.UTC(1899,11,30) + v * 86400000);
    return dt.getUTCFullYear()*10000 + (dt.getUTCMonth()+1)*100 + dt.getUTCDate();
  }
  var s = String(v).replace(/\D/g,'');
  if (s.length === 8) return parseInt(s,10);
  if (s.length === 7) return (parseInt(s.substring(0,3),10)+1911)*10000 + parseInt(s.substring(3),10); // 民國7碼
  return null;
}

/* 西元8碼 → 民國年 */
function ad8(v){
  var n = rawAd8(v);
  if (n === null) return null;
  return Math.floor(n / 10000) - 1911;
}

/* --- 分類分組類別 --- */
function catLabel(v){
  var s = String(v === undefined || v === null ? '' : v).trim();
  if (s === '1') return '一';
  if (s === '2') return '二';
  if (s === '3' || s === '4') return '三A';
  if (s === '5' || s === '6') return '三B';
  return s;
}

var BLANK = '＿＿＿＿＿＿';

function num(v){
  if (v === undefined || v === null || v === '') return null;
  var f = parseFloat(String(v).replace(/,/g,''));
  return isNaN(f) ? null : f;
}

function fmt(v, dec){
  var n = num(v);
  if (n === null) return '';
  return n.toLocaleString('en-US',{minimumFractionDigits:dec||0, maximumFractionDigits:dec===undefined?0:dec});
}

function money(v){
  var n = num(v);
  if (n === null) return '';
  return n.toLocaleString('en-US',{minimumFractionDigits:0, maximumFractionDigits:2});
}

function txt(v, fallback){
  if (v === undefined || v === null || String(v).trim() === '') return fallback === undefined ? '' : fallback;
  return String(v).trim();
}

function normPriceYear(v){
  var s = String(v === undefined || v === null ? '' : v).replace(/\D/g,'');
  return s;
}

/* 支付價年月 11508 → 「115年8月」 */
function pyLabel(py){
  var s = String(py === undefined || py === null ? '' : py).replace(/\D/g,'');
  if (s.length !== 5) return s;
  return parseInt(s.substring(0,3),10) + '年' + parseInt(s.substring(3,5),10) + '月';
}
