'use strict';
/* 四分頁 Excel 報表（分頁 1 以 skipPriceCalc 控制） */

function parseFilter(v){
  return String(v || '').trim().toLowerCase().split(/[\s,，、|]+/).filter(function(x){ return x; });
}

function passFilter(r, f){
  function m(list, val){
    if (!list.length) return true;
    val = String(val || '').toLowerCase();
    for (var i=0;i<list.length;i++) if (val.indexOf(list[i]) >= 0) return true;
    return false;
  }
  return m(f.atc7, atcOf(r.d)) && m(f.grpCd, r.grp) && m(f.grpNm, txt(r.d['分組名稱']))
      && m(f.code, r.code) && m(f.brand, txt(r.d['藥品名稱']));
}

/* opt：{ skipPriceCalc:Boolean }
   skipPriceCalc = true（路徑 A 背景查詢）→ 完全不產出分頁 1，也不觸碰 targets[i].calc，
   因此 price-calc.js / case-report.js 未載入時本函式仍可執行。
   預設 false（路徑 B 正式提案）→ 行為與單軌版本完全相同。 */
function buildSasWorkbook(model, filters, targets, priceYear, opt){
  var YR1 = model.YR1, YR2 = model.YR2, YR3 = model.YR3, priceCol = model.priceCol;
  var i, j;
  var skipPriceCalc = !!(opt && opt.skipPriceCalc);
  targets = targets || [];
  var sel = model.rows.filter(function(r){ return passFilter(r, filters); });

  /* ---------- 分頁1：藥價核定試算（skipPriceCalc 時整張略過） ---------- */
  var s1 = [], m1 = [], rh1 = {};
  function push1(row){ s1.push(row); return s1.length - 1; }
  if (!skipPriceCalc){
  push1([xcell('藥價核定試算', XS.TITLE)]);
  rh1[0] = 26;
  push1([xcell('依「全民健康保險藥物給付項目及支付標準」第35條特殊藥品支付價格訂定原則，'
    + '本案藥品就不同核價方式試算之支付價格及財務衝擊如下。（單位：元；財務衝擊為年度推估）', XS.NOTE)]);
  push1([]);
  for (i=0;i<targets.length;i++){
    var t = targets[i], c = t.calc;
    var n1 = c.tenHead.length, n2 = c.costHead.length, total = 2 + n1 + n2;
    var rSect = push1([xcell(txt(t.drug['藥品名稱']) + '（藥品代碼：' + t.code + '）', XS.SECT)]);

    var rHd = push1([xcell('健保價', XS.HEAD), xcell('廠商建議價', XS.HEAD), xcell('核價方式', XS.HEAD)]);
    m1.push(xlColName(2) + (rHd+1) + ':' + xlColName(total-1) + (rHd+1));

    var r2 = [xcell(t.row.price === null ? '－' : t.row.price, t.row.price === null ? XS.BODY : XS.DEC),
              xcell(t.sug === null ? '－' : t.sug, t.sug === null ? XS.BODY : XS.DEC),
              xcell('十國藥價中位數(註1)', XS.HEAD)];
    for (j=1;j<n1;j++) r2.push(xcell('', XS.HEAD));
    r2.push(xcell('參考成本價(註2)', XS.HEAD));
    for (j=1;j<n2;j++) r2.push(xcell('', XS.HEAD));
    var rSub = push1(r2);
    if (n1 > 1) m1.push(xlColName(2) + (rSub+1) + ':' + xlColName(2+n1-1) + (rSub+1));
    if (n2 > 1) m1.push(xlColName(2+n1) + (rSub+1) + ':' + xlColName(total-1) + (rSub+1));

    function calcRow(label, vals, numeric){
      var row = [xcell(label, XS.LABEL), xcell('', XS.LABEL)];
      for (var k=0;k<vals.length;k++){
        var v = vals[k];
        var asNum = numeric && v !== '無' && v !== '－' && v !== '' && !isNaN(Number(v));
        row.push(xcell(asNum ? Number(v) : v, asNum ? XS.DEC : XS.BODY));
      }
      var idx = push1(row);
      m1.push('A' + (idx+1) + ':B' + (idx+1));
    }
    var rPct = push1([xcell('加計比例', XS.LABEL), xcell('', XS.LABEL)]
      .concat(c.tenHead.map(function(x){ return xcell(x, XS.HEAD); }))
      .concat(c.costHead.map(function(x){ return xcell(x, XS.HEAD); })));
    m1.push('A' + (rPct+1) + ':B' + (rPct+1));
    calcRow('調高後支付價', c.priceRow, true);
    calcRow('整體藥費',     c.totalRow, false);
    calcRow('財務衝擊',     c.impactRow, false);

    push1([xcell('註1：' + c.a10Law, XS.NOTE)]);
    push1([xcell('註2：' + c.costLaw + '　2.領有藥物許可證者，得加計繳納藥害救濟徵收金比率及營業稅。' + c.licNote, XS.NOTE)]);
    push1([xcell('分組名稱：' + (t.grp ? t.grp.name : '')
      + '　　每月申報金額(' + YR3 + '年)：' + (!t.grp || t.grp.monthlyAmt === null ? '-' : fmt(Math.round(t.grp.monthlyAmt))) + ' 元'
      + '　　近三年平均申報量：' + (!t.grp || t.grp.avgQty3 === null ? '-' : fmt(Math.round(t.grp.avgQty3))), XS.NOTE)]);
    push1([]);
  }
  if (!targets.length) push1([xcell('未指定目標藥品CODE，無法計算適用級距。', XS.NOTE)]);
  push1([xcell('※ 整體藥費 ＝ 調高後支付價 × 同分組近三年平均申報量(QTY)', XS.NOTE)]);
  push1([xcell('※ 財務衝擊 ＝ (調高後支付價 － 健保價) × 同分組近三年平均申報量(QTY)', XS.NOTE)]);
  }

  var cols1 = [16, 14];
  for (i=2;i<24;i++) cols1.push(13);

  /* ---------- 分頁2：分組統計 ---------- */
  var S2H = ['分組代碼','分組名稱','調劑大類','[ATC代碼] ATC名稱','項目數','分組項目總數(含0元)',
             '支付價'+priceYear,
             YR1+'年申報金額分組總計', YR2+'年申報金額分組總計', YR3+'年申報金額分組總計','分組每月申報金額',
             YR1+'年申報數量分組總計', YR2+'年申報數量分組總計', YR3+'年申報數量分組總計','近三年平均申報量'];
  var ST2 = S2H.map(styleForHeader);
  ST2[6] = XS.BODY;                                     /* 支付價區間是文字 */
  var buckets = {}, bOrder = [];
  for (i=0;i<sel.length;i++){
    var key = sel[i].grp + '' + atcInfo(sel[i].d);
    if (!buckets[key]){ buckets[key] = {grp:sel[i].grp, atc:atcInfo(sel[i].d), rows:[]}; bOrder.push(key); }
    buckets[key].rows.push(sel[i]);
  }
  function statVals(label, rows, gcode){
    var g = model.groups[gcode];
    var amt = [0,0,0], qty = [0,0,0], items = [];
    for (var m=0;m<rows.length;m++){
      for (var n=0;n<3;n++){ amt[n] += rows[m].amtAdj[n]; qty[n] += rows[m].qtyAdj[n]; }
      items.push(rows[m].d);
    }
    /* 分組代碼留白者不是一個真的分組：不套用分組層彙總（支付價區間、近三年平均申報量） */
    var real = !!gcode;
    return [gcode || '(未填分組代碼)', real && g ? g.name : '', catLabel(rows[0].d['分類']), label,
            countValid(items, priceCol), rows.length, (real && g) ? g.priceRange : '',
            Math.round(amt[0]), Math.round(amt[1]), Math.round(amt[2]),
            Math.round(amt[2]/12),
            Math.round(qty[0]), Math.round(qty[1]), Math.round(qty[2]),
            (real && g && g.avgQty3 !== null) ? Math.round(g.avgQty3) : null];
  }
  var byGrp = {}, grpOrder = [];
  for (i=0;i<bOrder.length;i++){
    var bk = buckets[bOrder[i]];
    if (!byGrp[bk.grp]){ byGrp[bk.grp] = []; grpOrder.push(bk.grp); }
    byGrp[bk.grp].push(bk);
  }
  grpOrder.sort(function(a,b){
    var za = (model.groups[a] && model.groups[a].cnt === 0) ? 1 : 0;
    var zb = (model.groups[b] && model.groups[b].cnt === 0) ? 1 : 0;
    if (za !== zb) return za - zb;
    return a < b ? -1 : (a > b ? 1 : 0);
  });
  var s2 = [headRow(S2H)];
  var SUBST = S2H.map(function(h){ return styleForHeader(h) === XS.BODY ? XS.SUB : XS.SUBINT; });
  for (i=0;i<grpOrder.length;i++){
    var list = byGrp[grpOrder[i]].slice().sort(function(a,b){ return a.atc < b.atc ? -1 : 1; });
    var all = [];
    for (j=0;j<list.length;j++){
      s2.push(dataRow(statVals(list[j].atc, list[j].rows, grpOrder[i]), ST2));
      all = all.concat(list[j].rows);
    }
    if (list.length > 1) s2.push(dataRow(statVals('分組小計', all, grpOrder[i]), SUBST));
  }

  /* ---------- 分頁3 / 4：項目明細 ---------- */
  var wt = {};
  for (i=0;i<sel.length;i++){
    var wk = sel[i].grp + '' + sel[i].code.substring(0,8);
    wt[wk] = (wt[wk] || 0) + sel[i].amt[2];
  }
  var det = sel.slice().sort(function(a,b){
    var ca = catLabel(a.d['分類']), cb = catLabel(b.d['分類']);
    if (ca !== cb) return ca < cb ? -1 : 1;
    if (a.grp !== b.grp) return a.grp < b.grp ? -1 : 1;
    var za = (a.price === 0 ? 1 : 0), zb = (b.price === 0 ? 1 : 0);
    if (za !== zb) return za - zb;
    var wa = wt[a.grp+''+a.code.substring(0,8)] || 0;
    var wb2 = wt[b.grp+''+b.code.substring(0,8)] || 0;
    if (wa !== wb2) return wb2 - wa;
    if (a.amt[2] !== b.amt[2]) return b.amt[2] - a.amt[2];
    return a.code < b.code ? -1 : 1;
  });
  var targetCodes = {};
  for (i=0;i<targets.length;i++) targetCodes[targets[i].code] = 1;

  function detVals(r, withZero){
    var g = model.groups[r.grp];
    var amtShare = (g && g.amtSum[2]) ? adaptivePct(r.amt[2]/g.amtSum[2]*100) : '';
    var qtyShare = (g && g.qtySum[2]) ? adaptivePct(r.qty[2]/g.qtySum[2]*100) : '';
    var base = [
      r.grp, g ? g.name : '', catLabel(r.d['分類']), atcInfo(r.d), r.code,
      txt(r.d['藥品名稱']) + (targetCodes[r.code] ? '（本案藥品）' : ''),
      txt(r.d['藥商名稱']), txt(r.d['藥品分類_名稱']),
      r.price
    ];
    if (withZero) base.push(toRocDate(r.d['生效日期']));
    base = base.concat([
      Math.round(r.amt[0]), Math.round(r.amt[1]), Math.round(r.amt[2]), amtShare,
      Math.round(r.qty[0]), Math.round(r.qty[1]), Math.round(r.qty[2]), qtyShare,
      txt(r.d['適應症']), txt(r.d['藥價說明'])
    ]);
    if (withZero) base = base.concat([txt(r.d['不良品暫停支付註記']), num(r.d['不良品暫停支付前價格'])]);
    return base;
  }
  var DH = ['分組代碼','分組名稱','調劑大類','[ATC代碼] ATC名稱','CODE','藥品名稱','藥商名稱','藥品分類_名稱','支付價'+priceYear];
  var DT = [YR1+'年申報金額', YR2+'年申報金額', YR3+'年申報金額', YR3+'年AMT占率',
            YR1+'年申報數量', YR2+'年申報數量', YR3+'年申報數量', YR3+'年QTY占率', '適應症','藥價說明'];
  var H3 = DH.concat(DT);
  var H4 = DH.concat(['生效日期']).concat(DT).concat(['不良品暫停支付註記','不良品暫停支付前價格']);
  var ST3 = H3.map(styleForHeader), ST4 = H4.map(styleForHeader);
  ST3[12] = ST3[16] = XS.BODY;                          /* 占率是文字 */
  ST4[13] = ST4[17] = XS.BODY;

  var s3 = [headRow(H3)], s4 = [headRow(H4)];
  for (i=0;i<det.length;i++){
    var mk = targetCodes[det[i].code] ? {5: XS.MARK} : null;
    if (isCounted(det[i])) s3.push(dataRow(detVals(det[i], false), ST3, mk));
    s4.push(dataRow(detVals(det[i], true), ST4, mk));
  }

  function lastRef(aoa, headers){ return 'A1:' + xlColName(headers.length-1) + aoa.length; }

  var sheets = [];
  if (!skipPriceCalc)
    sheets.push({name:'1.藥價核定試算',  aoa:s1, cols:cols1, rowHeights:rh1, merges:m1, rowHeight:20});
  var n = skipPriceCalc ? 0 : 1;
  sheets.push({name:(n+1)+'.總覽篩選_分組統計',       aoa:s2, freeze:1, autoFilter:lastRef(s2,S2H), rowHeights:{0:34}, rowHeight:20});
  sheets.push({name:(n+2)+'.總覽篩選_項目明細',       aoa:s3, freeze:1, autoFilter:lastRef(s3,H3),  rowHeights:{0:34}, rowHeight:20});
  sheets.push({name:(n+3)+'.總覽篩選_項目明細(含0元)', aoa:s4, freeze:1, autoFilter:lastRef(s4,H4),  rowHeights:{0:34}, rowHeight:20});

  return {
    book: buildXlsxBook(sheets),
    counts: {s2:s2.length-1, s3:s3.length-1, s4:s4.length-1, sel:sel.length},
    skipPriceCalc: skipPriceCalc
  };
}
