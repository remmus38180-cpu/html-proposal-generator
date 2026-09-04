'use strict';
/* 路徑 A｜藥品背景查詢表（2 個工作表）
   輸入只有藥品主檔 + price 民國年月，完全不進行核價試算，
   因此本檔不依賴 price-calc.js / case-report.js / docx-*。

   Sheet 1 同分組項目明細    先依篩選條件定位出目標分組，再展開該分組旗下「全部品項」
                              （含 0 元項目），並標記命中篩選條件的品項；
                              增加「含量」「規格量」欄，因同分組可能有不同含量或規格。
   Sheet 2 分組與 ATC 彙總  同分組代碼 × ATC7 碼 聚合層級；
                              顯示每個組合的項目數與三年申報量／金額加總。

   簡化規則（INTENT 路徑 A）：不出現廠商建議價、十國中位價、廠商成本、
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
  var YR1 = model.YR1, YR2 = model.YR2, YR3 = model.YR3;
  var i, j, k;
  /* sel：直接符合篩選條件的項目，僅用來定位分組代碼 */
  var sel = queryScope(model, filters);
  var stamp = queryStamp();

  function titleNote(extra){
    return QUERY_BANNER + '　／　查詢日期 ' + stamp
      + '　／　支付價基準 ' + pyLabel(priceYear)
      + '　／　申報年度視窗 民國 ' + YR1 + '–' + YR3 + ' 年'
      + (extra ? '　／　' + extra : '');
  }
  function footRows(){
    return [[], [xcell('※ ' + QUERY_BANNER + '，不含任何核價試算欄位。', XS.NOTE)]];
  }

  /* ── 由篩選命中項目的分組代碼，收集同分組全部項目（含 0 元） ── */
  var seenG = {}, gList = [];
  for (i=0;i<sel.length;i++){
    if (!sel[i].grp || seenG[sel[i].grp]) continue;
    seenG[sel[i].grp] = 1; gList.push(sel[i].grp);
  }
  /* queryCodes：命中篩選條件的 code 集合，用來標記 Sheet 1 的「本查詢項目」 */
  var queryCodes = {};
  for (i=0;i<sel.length;i++) queryCodes[sel[i].code] = 1;

  /* 蒐集同分組全部項目（所有分組的 rows 合併） */
  var grpRows = [];
  for (k=0;k<gList.length;k++){
    var gg = model.groups[gList[k]];
    if (!gg) continue;
    for (i=0;i<gg.rows.length;i++) grpRows.push(gg.rows[i]);
  }

  /* ---------- Sheet 1：同分組項目明細 ----------
     粒度：品項（CODE）層級，含全部品項（含 0 元）
     欄位：分組代碼(A) CODE(B) 中文名稱(C) 藥品名稱(D) ATC7碼(E) ATC名稱(F)
           支付價(G) 含量(H) 規格量(I) 分組名稱(J) 劑型(K) 藥商名稱(L)
           YR1-3 申報量(M-R) YR1-3 申報金額  ← 數量在前，金額在後（同分頁慣例）
     排序：調劑大類↑ → 分組代碼↑ → 0元置後 → 許可證層AMT↓ → 當期AMT↓ → CODE↑
     標記：命中篩選條件的品項以淡藍底色標出（同提案附件「本案藥品」慣例）
     小計：每個分組結束時插入灰底小計列（含未列出項目的加總）             */
  var H1 = ['分組代碼', 'CODE', '中文名稱', '藥品名稱', 'ATC7碼', 'ATC名稱',
            '支付價' + priceYear, '含量', '規格量',
            '分組名稱', '劑型', '藥商名稱',
            YR1 + '年申報量', YR1 + '年申報金額',
            YR2 + '年申報量', YR2 + '年申報金額',
            YR3 + '年申報量', YR3 + '年申報金額'];
  var NCOL1 = H1.length;
  var ST1 = H1.map(styleForHeader);
  ST1[0] = ST1[1] = ST1[4] = XS.BODY;   /* 分組代碼/CODE/ATC7碼 文字 */
  ST1[7] = ST1[8] = XS.BODY;            /* 含量/規格量 可含單位，保留文字 */

  /* 排序 */
  var wt1 = {};
  for (i=0;i<grpRows.length;i++){
    var wk = grpRows[i].grp + '' + grpRows[i].code.substring(0,8);
    wt1[wk] = (wt1[wk] || 0) + grpRows[i].amt[2];
  }
  grpRows.sort(function(a,b){
    var na = a.grp ? 0 : 1, nb = b.grp ? 0 : 1;
    if (na !== nb) return na - nb;
    var ca = catLabel(a.d['分類']), cb = catLabel(b.d['分類']);
    if (ca !== cb) return ca < cb ? -1 : 1;
    if (a.grp !== b.grp) return a.grp < b.grp ? -1 : 1;
    var za = (a.price === 0 ? 1 : 0), zb = (b.price === 0 ? 1 : 0);
    if (za !== zb) return za - zb;
    var wa = wt1[a.grp+''+a.code.substring(0,8)] || 0;
    var wb2 = wt1[b.grp+''+b.code.substring(0,8)] || 0;
    if (wa !== wb2) return wb2 - wa;
    if (a.amt[2] !== b.amt[2]) return b.amt[2] - a.amt[2];
    return a.code < b.code ? -1 : 1;
  });

  var s1 = [
    [xcell('同分組項目明細', XS.TITLE)],
    [xcell(titleNote('先以篩選條件定位分組，再展開該分組旗下全部品項（含 0 元項目）'
      + '　※淡藍底色＝命中篩選條件的品項　※灰底＝分組小計（含未列出項目）'), XS.NOTE)],
    H1.map(function(h){ return xcell(h, XS.HEAD); })
  ];

  var cur1 = null, gAmt1 = [0,0,0], gQty1 = [0,0,0], tAmt1 = [0,0,0], tQty1 = [0,0,0];
  var nItems1 = 0, nGrps1 = 0;

  function subtotal1(gcode){
    var g = model.groups[gcode];
    var row = new Array(NCOL1);
    for (var k=0;k<NCOL1;k++) row[k] = xcell('', XS.SUB);
    row[0]  = xcell(gcode, XS.SUB);
    row[3]  = xcell('分組小計', XS.SUB);
    row[9]  = xcell(g ? g.name : '', XS.SUB);
    row[12] = xcell(Math.round(gQty1[0]), XS.SUBINT);
    row[13] = xcell(Math.round(gAmt1[0]), XS.SUBINT);
    row[14] = xcell(Math.round(gQty1[1]), XS.SUBINT);
    row[15] = xcell(Math.round(gAmt1[1]), XS.SUBINT);
    row[16] = xcell(Math.round(gQty1[2]), XS.SUBINT);
    row[17] = xcell(Math.round(gAmt1[2]), XS.SUBINT);
    return row;
  }

  for (i=0;i<grpRows.length;i++){
    var r = grpRows[i], d = r.d;
    if (cur1 !== null && r.grp !== cur1){
      s1.push(subtotal1(cur1));
      nGrps1++;
      gAmt1 = [0,0,0]; gQty1 = [0,0,0];
    }
    cur1 = r.grp;
    var vals1 = [
      r.grp, r.code,
      txt(d['中文名稱']), txt(d['藥品名稱']),
      atcOf(d), txt(d['ATC名稱']),
      (r.price === 0 ? '－' : r.price),           /* 0元不顯示數字 */
      txt(d['成分含量']) || txt(d['成分及含量']),  /* 含量 */
      txt(d['規格量']),                             /* 規格量 */
      txt(d['分組名稱']), txt(d['劑型']), txt(d['藥商名稱']),
      Math.round(r.qtyAdj[0]), Math.round(r.amtAdj[0]),
      Math.round(r.qtyAdj[1]), Math.round(r.amtAdj[1]),
      Math.round(r.qtyAdj[2]), Math.round(r.amtAdj[2])
    ];
    s1.push(dataRow(vals1, ST1, queryCodes[r.code] ? {3: XS.MARK} : null));
    nItems1++;
    for (j=0;j<3;j++){ gQty1[j] += r.qtyAdj[j]; gAmt1[j] += r.amtAdj[j]; tQty1[j] += r.qtyAdj[j]; tAmt1[j] += r.amtAdj[j]; }
  }
  if (cur1 !== null){ s1.push(subtotal1(cur1)); nGrps1++; }

  /* 總計列 */
  if (nItems1){
    var tot1 = new Array(NCOL1);
    for (i=0;i<NCOL1;i++) tot1[i] = xcell('', XS.SUB);
    tot1[3]  = xcell('總計', XS.SUB);
    tot1[12] = xcell(Math.round(tQty1[0]), XS.SUBINT); tot1[13] = xcell(Math.round(tAmt1[0]), XS.SUBINT);
    tot1[14] = xcell(Math.round(tQty1[1]), XS.SUBINT); tot1[15] = xcell(Math.round(tAmt1[1]), XS.SUBINT);
    tot1[16] = xcell(Math.round(tQty1[2]), XS.SUBINT); tot1[17] = xcell(Math.round(tAmt1[2]), XS.SUBINT);
    s1.push(tot1);
  }
  if (!nItems1) s1.push([xcell('查無符合篩選條件的分組項目。', XS.NOTE)]);
  s1 = s1.concat(footRows());

  /* ---------- Sheet 2：分組與 ATC 彙總 ----------
     粒度：分組代碼 × ATC7碼 聚合層級
     資料來源：同 Sheet 1 的分組範圍，但遵守「計數規則」：
               排除「支付價＝0 且不良品暫停支付註記≠Y」的品項（isCounted）
     欄位：分組代碼(A) 分組名稱(B) ATC7碼(C) ATC名稱(D) 項目數(E)
           YR1-3 申報量(F-K) YR1-3 申報金額
     項目數：該分組×ATC 組合的有效品項數（已排除 0 元且非不良品者）
     排序：分組代碼↑ → ATC7碼↑
     小計：每個分組結束時插入灰底小計列（該分組全部 ATC 加總）         */
  var grpAtcMap = {}, grpAtcOrder = [];
  for (i=0;i<grpRows.length;i++){
    var r2 = grpRows[i], d2 = r2.d;
    if (!isCounted(r2)) continue;   /* 排除 0 元且不良品註記≠Y 的品項 */
    var gc2 = r2.grp, atc2 = atcOf(d2);
    var key2 = gc2 + '|' + atc2;
    if (!grpAtcMap[key2]){
      grpAtcMap[key2] = {
        grp: gc2, grpName: txt(d2['分組名稱']),
        atc: atc2, atcName: txt(d2['ATC名稱']),
        cnt: 0, qty: [0,0,0], amt: [0,0,0]
      };
      grpAtcOrder.push(key2);
    }
    var ent = grpAtcMap[key2];
    ent.cnt++;
    for (j=0;j<3;j++){ ent.qty[j] += r2.qtyAdj[j]; ent.amt[j] += r2.amtAdj[j]; }
  }
  grpAtcOrder.sort(function(a, b){
    var ea = grpAtcMap[a], eb = grpAtcMap[b];
    if (ea.grp !== eb.grp) return ea.grp < eb.grp ? -1 : 1;
    return ea.atc < eb.atc ? -1 : (ea.atc > eb.atc ? 1 : 0);
  });

  var H2 = ['分組代碼', '分組名稱', 'ATC7碼', 'ATC名稱', '項目數',
            YR1 + '年申報量', YR1 + '年申報金額',
            YR2 + '年申報量', YR2 + '年申報金額',
            YR3 + '年申報量', YR3 + '年申報金額'];
  var NCOL2 = H2.length;
  var ST2 = H2.map(styleForHeader);
  ST2[0] = ST2[2] = XS.BODY;   /* 分組代碼/ATC7碼 文字 */

  var s2 = [
    [xcell('分組與ATC彙總', XS.TITLE)],
    [xcell(titleNote('聚合層級：分組代碼 × ATC7碼'
      + '　※項目數排除「0元且不良品註記≠Y」品項　※灰底＝分組小計與總計'), XS.NOTE)],
    H2.map(function(h){ return xcell(h, XS.HEAD); })
  ];

  var cur2 = null, cur2Nm = '', gCnt2 = 0, gQty2 = [0,0,0], gAmt2 = [0,0,0];
  var tCnt2 = 0, tQty2 = [0,0,0], tAmt2 = [0,0,0];
  var nRows2 = 0;

  function subtotal2(gcode, gnm){
    var row = new Array(NCOL2);
    for (var k=0;k<NCOL2;k++) row[k] = xcell('', XS.SUB);
    row[0] = xcell(gcode, XS.SUB);
    row[1] = xcell(gnm,   XS.SUB);
    row[2] = xcell('分組小計', XS.SUB);
    row[4] = xcell(gCnt2, XS.SUBINT);
    row[5] = xcell(Math.round(gQty2[0]), XS.SUBINT); row[6]  = xcell(Math.round(gAmt2[0]), XS.SUBINT);
    row[7] = xcell(Math.round(gQty2[1]), XS.SUBINT); row[8]  = xcell(Math.round(gAmt2[1]), XS.SUBINT);
    row[9] = xcell(Math.round(gQty2[2]), XS.SUBINT); row[10] = xcell(Math.round(gAmt2[2]), XS.SUBINT);
    return row;
  }

  /* 每個分組有幾種 ATC，單一 ATC 的分組不輸出小計 */
  var grpAtcCount = {};
  for (i=0;i<grpAtcOrder.length;i++){
    var gc_ = grpAtcMap[grpAtcOrder[i]].grp;
    grpAtcCount[gc_] = (grpAtcCount[gc_] || 0) + 1;
  }

  for (i=0;i<grpAtcOrder.length;i++){
    var ent2 = grpAtcMap[grpAtcOrder[i]];
    if (cur2 !== null && ent2.grp !== cur2){
      if (grpAtcCount[cur2] > 1) s2.push(subtotal2(cur2, cur2Nm));
      gCnt2 = 0; gQty2 = [0,0,0]; gAmt2 = [0,0,0];
    }
    cur2 = ent2.grp; cur2Nm = ent2.grpName;
    var vals2 = [
      ent2.grp, ent2.grpName, ent2.atc, ent2.atcName, ent2.cnt,
      Math.round(ent2.qty[0]), Math.round(ent2.amt[0]),
      Math.round(ent2.qty[1]), Math.round(ent2.amt[1]),
      Math.round(ent2.qty[2]), Math.round(ent2.amt[2])
    ];
    s2.push(dataRow(vals2, ST2, null));
    nRows2++;
    gCnt2 += ent2.cnt; tCnt2 += ent2.cnt;
    for (j=0;j<3;j++){ gQty2[j] += ent2.qty[j]; gAmt2[j] += ent2.amt[j]; tQty2[j] += ent2.qty[j]; tAmt2[j] += ent2.amt[j]; }
  }
  if (cur2 !== null && grpAtcCount[cur2] > 1) s2.push(subtotal2(cur2, cur2Nm));
  if (nRows2){
    var tot2 = new Array(NCOL2);
    for (i=0;i<NCOL2;i++) tot2[i] = xcell('', XS.SUB);
    tot2[2]  = xcell('總計', XS.SUB);
    tot2[4]  = xcell(tCnt2, XS.SUBINT);
    tot2[5]  = xcell(Math.round(tQty2[0]), XS.SUBINT); tot2[6]  = xcell(Math.round(tAmt2[0]), XS.SUBINT);
    tot2[7]  = xcell(Math.round(tQty2[1]), XS.SUBINT); tot2[8]  = xcell(Math.round(tAmt2[1]), XS.SUBINT);
    tot2[9]  = xcell(Math.round(tQty2[2]), XS.SUBINT); tot2[10] = xcell(Math.round(tAmt2[2]), XS.SUBINT);
    s2.push(tot2);
  }
  if (!nRows2) s2.push([xcell('查無符合篩選條件的分組。', XS.NOTE)]);
  s2 = s2.concat(footRows());

  return {
    book: buildXlsxBook([
      {name:'1.同分組項目明細', aoa:s1, cols:autoCols(s1, 10, 42), freeze:3, rowHeights:{0:26,2:34}, rowHeight:20},
      {name:'2.分組與ATC彙總', aoa:s2, cols:autoCols(s2, 10, 42), freeze:3, rowHeights:{0:26,2:34}, rowHeight:20}
    ]),
    counts: {sel:sel.length, groups:nGrps1, atcRows:nRows2,
             s1:nItems1, s2:nRows2}
  };
}
