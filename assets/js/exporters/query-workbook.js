'use strict';
/* 路徑 A｜藥品背景查詢表（3 個工作表）
   輸入只有藥品主檔 + price 民國年月，完全不進行核價試算，
   因此本檔不依賴 price-calc.js / case-report.js / docx-*。

   Sheet 1 查詢摘要      符合篩選條件的項目本身
   Sheet 2 同分組比較    上述項目所屬分組的全部項目（分組層彙總小計）
   Sheet 3 同ATC相關品項 上述項目 ATC7 碼相同的全部項目（可能落在不同分組）

   簡化規則（INTENT 路徑 A）：不出現 PRICE、廠商建議價、十國中位價、廠商成本、
   調高後支付價、財務衝擊、核定備註；保留分組層彙總（支付價區間、項目數、申報量）。 */

var QUERY_BANNER = '此為背景查詢版，非正式提案';

function queryStamp(){
  var d = new Date();
  var p = function(n){ return (n < 10 ? '0' : '') + n; };
  return d.getFullYear() + '-' + p(d.getMonth()+1) + '-' + p(d.getDate());
}

/* 查詢範圍：套用步驟 03 的篩選條件；留空 → 全主檔 */
function queryScope(model, filters){
  return model.rows.filter(function(r){ return passFilter(r, filters); });
}

function buildQueryWorkbook(model, filters, priceYear){
  var YR1 = model.YR1, YR2 = model.YR2, YR3 = model.YR3, priceCol = model.priceCol;
  var i, k;
  var sel = queryScope(model, filters);
  var stamp = queryStamp();

  function titleRows(title, note){
    return [
      [xcell(title, XS.TITLE)],
      [xcell(QUERY_BANNER + '　／　查詢日期 ' + stamp
        + '　／　支付價基準 ' + pyLabel(priceYear)
        + '　／　申報年度視窗 民國 ' + YR1 + '–' + YR3 + ' 年'
        + (note ? '　／　' + note : ''), XS.NOTE)],
      []
    ];
  }
  function footRows(){
    return [[], [xcell('※ ' + QUERY_BANNER + '，不含任何核價試算欄位。', XS.NOTE)]];
  }
  function grpOf(r){ return model.groups[r.grp]; }

  /* ---------- Sheet 1：查詢摘要 ---------- */
  var H1 = ['CODE','藥品名稱','藥商名稱','ATC7碼','ATC名稱','分組代碼','分組名稱',
            '收載日期','生效日期','藥品分類_名稱','必要藥品',
            '同分組項目數','分組項目總數(含0元)','支付價區間'+priceYear,
            '同成分同劑型收載年','第三B大類'];
  var ST1 = H1.map(styleForHeader);
  ST1[13] = XS.BODY;                          /* 支付價區間是文字 */
  var s1 = titleRows('藥品背景查詢　·　查詢摘要');
  s1.push(headRow(H1));
  for (i=0;i<sel.length;i++){
    var r = sel[i], g = grpOf(r);
    s1.push(dataRow([
      r.code, txt(r.d['藥品名稱']), txt(r.d['藥商名稱']),
      atcOf(r.d), txt(r.d['ATC名稱']).toLowerCase(),
      r.grp, g ? g.name : '',
      toRocDate(r.d['收載日期']), toRocDate(r.d['生效日期']),
      drugClassLabel(r.d), essentialLabel(r.d['必要藥品']),
      g ? g.cnt : null, g ? g.total : null, g ? g.priceRange : '',
      (r.ingListYear === null || r.ingListYear === undefined) ? '' : r.ingListYear,
      r.cat3B === null || r.cat3B === undefined ? '' : (r.cat3B ? 'Y' : 'N')
    ], ST1));
  }
  if (!sel.length) s1.push([xcell('查無符合篩選條件的項目。', XS.NOTE)]);
  s1 = s1.concat(footRows());

  /* ---------- Sheet 2：同分組比較 ---------- */
  var H2 = ['分組代碼','分組名稱','調劑大類','[ATC代碼] ATC名稱','CODE','藥品名稱','藥商名稱',
            '藥品分類_名稱','收載日期',
            YR1+'年申報數量', YR2+'年申報數量', YR3+'年申報數量',
            YR1+'年申報金額', YR2+'年申報金額', YR3+'年申報金額'];
  var ST2 = H2.map(styleForHeader);
  var seenG = {}, gList = [];
  for (i=0;i<sel.length;i++){
    if (!sel[i].grp || seenG[sel[i].grp]) continue;
    seenG[sel[i].grp] = 1; gList.push(sel[i].grp);
  }
  var s2 = titleRows('藥品背景查詢　·　同分組比較',
                     '分組層金額／數量為整組全部項目（含未列出者）之加總');
  s2.push(headRow(H2));
  for (k=0;k<gList.length;k++){
    var gc = gList[k], gg = model.groups[gc];
    if (!gg) continue;
    for (i=0;i<gg.rows.length;i++){
      var rr = gg.rows[i];
      if (!isCounted(rr)) continue;                     /* [B5] 給付中項目才列出 */
      s2.push(dataRow([
        gc, gg.name, catLabel(rr.d['分類']), atcInfo(rr.d), rr.code,
        txt(rr.d['藥品名稱']), txt(rr.d['藥商名稱']), drugClassLabel(rr.d),
        toRocDate(rr.d['收載日期']),
        Math.round(rr.qtyAdj[0]), Math.round(rr.qtyAdj[1]), Math.round(rr.qtyAdj[2]),
        Math.round(rr.amtAdj[0]), Math.round(rr.amtAdj[1]), Math.round(rr.amtAdj[2])
      ], ST2));
    }
    /* 分組小計：整組（含未列出項目）之加總，與 [B6]/[B7] 同一份數字 */
    s2.push(dataRow([
      gc, gg.name + '　小計', '', '項目數 ' + gg.cnt + ' ／ 總數 ' + gg.total,
      '支付價區間：' + gg.priceRange, '近三年平均申報量 ' + fmt(Math.round(gg.avgQty3)),
      '每月申報金額 ' + fmt(Math.round(gg.monthlyAmt)) + ' 元', '', '',
      Math.round(gg.qtySum[0]), Math.round(gg.qtySum[1]), Math.round(gg.qtySum[2]),
      Math.round(gg.amtSum[0]), Math.round(gg.amtSum[1]), Math.round(gg.amtSum[2])
    ], ST2, {0:XS.SUB,1:XS.SUB,2:XS.SUB,3:XS.SUB,4:XS.SUB,5:XS.SUB,6:XS.SUB,7:XS.SUB,8:XS.SUB,
             9:XS.SUBINT,10:XS.SUBINT,11:XS.SUBINT,12:XS.SUBINT,13:XS.SUBINT,14:XS.SUBINT}));
    s2.push([]);
  }
  if (!gList.length) s2.push([xcell('查無符合篩選條件的分組。', XS.NOTE)]);
  s2 = s2.concat(footRows());

  /* ---------- Sheet 3：同 ATC7 碼相關品項 ---------- */
  var H3 = ['ATC7碼','ATC名稱','CODE','藥品名稱','藥商名稱','分組代碼','分組名稱',
            '藥品分類_名稱','收載日期',
            YR3+'年申報數量', YR3+'年申報金額','同分組近三年平均申報量'];
  var ST3 = H3.map(styleForHeader);
  var seenA = {}, aList = [];
  for (i=0;i<sel.length;i++){
    var a = atcOf(sel[i].d);
    if (!a || seenA[a]) continue;
    seenA[a] = 1; aList.push(a);
  }
  var byAtc = {};
  for (i=0;i<model.rows.length;i++){
    var av = atcOf(model.rows[i].d);
    if (!seenA[av]) continue;
    (byAtc[av] = byAtc[av] || []).push(model.rows[i]);
  }
  var s3 = titleRows('藥品背景查詢　·　同 ATC7 碼相關品項',
                     '同一 ATC7 碼可能橫跨不同分組');
  s3.push(headRow(H3));
  for (k=0;k<aList.length;k++){
    var list = byAtc[aList[k]] || [];
    for (i=0;i<list.length;i++){
      var t = list[i], tg = model.groups[t.grp];
      if (!isCounted(t)) continue;
      s3.push(dataRow([
        atcOf(t.d), txt(t.d['ATC名稱']).toLowerCase(), t.code,
        txt(t.d['藥品名稱']), txt(t.d['藥商名稱']), t.grp, tg ? tg.name : '',
        drugClassLabel(t.d), toRocDate(t.d['收載日期']),
        Math.round(t.qtyAdj[2]), Math.round(t.amtAdj[2]),
        tg ? Math.round(tg.avgQty3) : null
      ], ST3));
    }
  }
  if (!aList.length) s3.push([xcell('查無符合篩選條件的 ATC7 碼。', XS.NOTE)]);
  s3 = s3.concat(footRows());

  function ref(aoa, headers){ return 'A4:' + xlColName(headers.length-1) + aoa.length; }

  return {
    book: buildXlsxBook([
      {name:'1.查詢摘要',      aoa:s1, cols:autoCols(s1, 10, 42), freeze:4, autoFilter:ref(s1,H1), rowHeights:{0:26,3:34}, rowHeight:20},
      {name:'2.同分組比較',    aoa:s2, cols:autoCols(s2, 10, 42), freeze:4, rowHeights:{0:26,3:34}, rowHeight:20},
      {name:'3.同ATC相關品項', aoa:s3, cols:autoCols(s3, 10, 42), freeze:4, autoFilter:ref(s3,H3), rowHeights:{0:26,3:34}, rowHeight:20}
    ]),
    counts: {sel:sel.length, groups:gList.length, atc:aList.length,
             s1:Math.max(s1.length-6,0), s2:Math.max(s2.length-6,0), s3:Math.max(s3.length-6,0)}
  };
}
