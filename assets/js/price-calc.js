'use strict';
/* 核價試算：[C9]~[C14]（僅 Path B 載入） */

/* --- [C12] 支付標準第40條：核算價小數處理（一律無條件捨去）
       截斷前先依 [C11] 消除二進位浮點誤差（truncatePrice 內含） --- */
function article40RoundDown(v){
  if (v === null || v === undefined || isNaN(v)) return null;
  return truncatePrice(v);
}

/* --- 參考十國藥價法核算價：上限價 = 中位數 × (1 + 加成率) --- */
function a10RefPrice(ten, pct){
  if (ten === null) return null;
  return article40RoundDown(ten * (1 + pct));
}

/* --- [C10] 參考成本價核算價
       1. 廠商成本 × (1 + 管銷費用比例)
       2. [C11] 消除浮點誤差後取小數兩位，第三位（含）以後捨去
       3. 領有藥物許可證 = Y → × 1.0505（營業稅 5% + 藥害救濟徵收金 0.05%）
       4. 依 [C12] 第40條截斷 --- */
function costRefPrice(cost, pct, lic){
  if (cost === null) return null;
  var base = truncate2(cost * (1 + pct));
  var fin = (String(lic || '').toUpperCase() === 'Y') ? base * (1 + 0.05 + 0.0005) : base;
  return article40RoundDown(fin);
}

/* --- 金額級距（依法條：≤50萬 / >50萬且≤100萬 / >100萬） --- */
var A10_PCTS  = [0.00,0.05,0.10,0.15,0.20];

var COST_PCTS = [0.00,0.05,0.10,0.15,0.20,0.25,0.30,0.35,0.40,0.45,0.50];

var A10_TIER_PCT  = [0.20, 0.10, 0.00];

var COST_TIER_PCT = [0.50, 0.40, 0.30];

var A10_TIER_LAW = [
  '（二）1.每月申報金額小於等於新臺幣五十萬元者，以十國藥價中位數加百分之二十為上限價。',
  '（二）2.每月申報金額大於新臺幣五十萬元、小於等於新臺幣一百萬元者，以十國藥價中位數加百分之十為上限價。',
  '（二）3.每月申報金額大於新臺幣一百萬元者，以十國藥價中位數為上限價。'
];

var COST_TIER_LAW = [
  '（1）每月申報金額小於等於新臺幣五十萬元者，加計百分之五十。',
  '（2）每月申報金額大於新臺幣五十萬元、小於等於新臺幣一百萬元者，加計百分之四十。',
  '（3）每月申報金額大於新臺幣一百萬元者，加計百分之三十。'
];

var LIC_NOTE = {
  Y: '因領有藥物許可證者，得加計繳納藥害救濟徵收金比率0.05%及營業稅5%。',
  N: '因未領有藥物許可證者，不得加計繳納藥害救濟徵收金比率0.05%及營業稅5%。'
};

/* 依法條文字「小於等於」判定（SAS 原碼用 low -< 500000，恰好 50 萬會落到第2級，此處已修正） */
function amtTier(monthlyAmt){
  if (monthlyAmt === null || monthlyAmt === undefined || isNaN(monthlyAmt)) return null;
  monthlyAmt = fpClean(monthlyAmt);
  if (monthlyAmt <= 500000)  return 0;
  if (monthlyAmt <= 1000000) return 1;
  return 2;
}

/* --- 第40條適用級距說明 --- */
function article40Text(lt5, l550, ge50){
  var s = '※ 全民健康保險藥物給付項目及支付標準第40條：核算價格小數點之處理方式如下（僅列本次實際適用之級距）：';
  if (lt5)  s += '\n一、核算價小於新臺幣5元者，取至小數點後第2位，第3位（含）以後無條件捨去。';
  if (l550) s += '\n二、核算價大於或等於新臺幣5元且小於50元者，取至小數點後第1位，第2位（含）以後無條件捨去。';
  if (ge50) s += '\n三、核算價大於或等於新臺幣五十元者，取至整數，小數點以後無條件捨去。';
  return s;
}

var SHEET1_ROWS = ['核算價','整體藥費','財務影響'];

function buildPriceCalc(opt){
  // opt: {nhi, sug, ten, cost, lic, avgQty, monthlyAmt}
  var ten = opt.ten, cost = opt.cost, avg = opt.avgQty;
  var tenV  = A10_PCTS.map(function(p){ return a10RefPrice(ten, p); });
  var costV = COST_PCTS.map(function(p){ return costRefPrice(cost, p, opt.lic); });
  var all = tenV.concat(costV);
  var lt5 = 0, l550 = 0, ge50 = 0;
  for (var i=0;i<all.length;i++){
    var v = all[i];
    if (v === null) continue;
    if (v < 5) lt5 = 1; else if (v < 50) l550 = 1; else ge50 = 1;
  }
  function fmtP(v){ return v === null ? '' : String(v); }
  function overall(v){ return (v === null || avg === null) ? '' : chineseAmount(v * avg); }
  function impact(v){ return (v === null || avg === null || opt.nhi === null) ? '' : chineseAmount((v - opt.nhi) * avg); }

  var body = [
    [ opt.nhi === null ? '' : String(opt.nhi), opt.sug === null ? '' : String(opt.sug), '核算價' ]
      .concat(tenV.map(fmtP)).concat(costV.map(fmtP)),
    ['-','-','整體藥費'].concat(tenV.map(overall)).concat(costV.map(overall)),
    ['-','-','財務影響'].concat(tenV.map(impact)).concat(costV.map(impact))
  ];
  var tier = amtTier(opt.monthlyAmt);
  return {
    tenVals: tenV, costVals: costV, rows: body, tier: tier,
    a10Pct:  tier === null ? null : A10_TIER_PCT[tier],
    costPct: tier === null ? null : COST_TIER_PCT[tier],
    a10Law:  tier === null ? '' : A10_TIER_LAW[tier],
    costLaw: tier === null ? '' : COST_TIER_LAW[tier],
    licNote: LIC_NOTE[String(opt.lic || '').toUpperCase()] || '',
    art40:   article40Text(lt5, l550, ge50)
  };
}

function priceCalcHeaders(priceYear){
  var h1 = ['健保支付價' + priceYear + '(元)', '廠商建議價(元)', ''];
  var h2 = A10_PCTS.map(function(p){ return Math.round(p*100) + '%'; });
  var h3 = COST_PCTS.map(function(p){ return Math.round(p*100) + '%'; });
  return {fixed:h1, ten:h2, cost:h3};
}

/* 「0,10,15」→ [0,10,15]；空字串 → [] */
function parsePcts(v){
  return String(v === undefined || v === null ? '' : v)
    .split(/[^0-9.]+/).filter(function(x){ return x !== ''; })
    .map(Number).filter(function(x){ return !isNaN(x); });
}
