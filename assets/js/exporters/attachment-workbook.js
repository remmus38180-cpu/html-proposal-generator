'use strict';
/* 提案附件 Excel */

var ATT_HEAD_FIXED = ['分組代碼','項目','ATC7碼','ATC名稱','CODE','中文名稱','藥品名稱'];

/* fixedRow=true：所有資料列鎖定固定列高，不隨 適應症 等長文字自動變高（列數極多的分頁用） */
function attachSheet(name, title, scopeRows, model, priceYear, caseCodes, fixedRow){
  var YR1 = model.YR1, YR2 = model.YR2, YR3 = model.YR3, priceCol = model.priceCol;
  var i, j;
  /* scopeRows＝本分頁的「納入範圍」，含依列出規則不會印出的項目（小計要涵蓋它們）。
     列出規則（§3.3.1 規則 B）：支付價 ≠ 0；或支付價 ＝ 0 且 不良品暫停支付註記 = 'Y'。
     申報量／金額一律採 qtyAdj／amtAdj（規則 A：依支付價與生效日期做年度歸零）。 */
  var sel = scopeRows.slice();

  /* 排序：調劑大類↑ → 分組代碼↑ → 0元置後 → 許可證層級加總AMT↓ → 當期AMT↓ → CODE↑ */
  var wt = {};
  for (i=0;i<sel.length;i++){
    var wk = sel[i].grp + '' + sel[i].code.substring(0,8);
    wt[wk] = (wt[wk] || 0) + sel[i].amt[2];
  }
  sel.sort(function(a,b){
    var na = a.grp ? 0 : 1, nb = b.grp ? 0 : 1;      /* 未編分組代碼者排最後 */
    if (na !== nb) return na - nb;
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

  var H = ATT_HEAD_FIXED.concat([
    '支付價' + priceYear, '藥品分類_名稱', '必要藥品', '分組名稱', '劑型', '適應症', '藥商名稱',
    YR1+'年申報量', YR1+'年申報金額', YR2+'年申報量', YR2+'年申報金額', YR3+'年申報量', YR3+'年申報金額'
  ]);
  var NCOL = H.length;                         /* 20 欄：A…T */
  var ST = H.map(styleForHeader);
  ST[0] = ST[2] = ST[4] = XS.BODY;             /* 分組代碼／ATC7碼／CODE 為文字 */
  ST[8] = ST[9] = XS.BODY;                     /* 藥品分類_名稱／必要藥品 */

  /* 標題與說明置於 B 欄，與列印範圍（B 起）對齊 */
  var aoa = [
    ['', xcell(title, XS.TITLE)],
    ['', xcell('※本表只列出支付價≠0，或支付價=0且不良品暫停支付註記=Y 的項目　'
         + '※申報量／金額已依支付價與生效日期做年度歸零　'
         + '※灰底小計為本表納入範圍內該分組的加總（含未列出之0元項目）　'
         + '※A欄「分組代碼」僅供審查，已排除於列印範圍　'
         + '※支付價欄位為 ' + pyLabel(priceYear) + ' 健保支付價', XS.NOTE)],
    H.map(function(h){ return xcell(h, XS.HEAD); })
  ];
  var seq = 0, gAmt = [0,0,0], gQty = [0,0,0], cur = null, buf = [];

  /* 分組小計＝「本分頁納入範圍內」該分組的全部項目加總（含依列出規則未印出者）。
     同分組分頁的範圍就是整組，因此其小計＝Word 表二的「整組申報金額／申報量」。
     ATC 分頁的範圍是 ATC 相符的項目，小計即該分組在本頁的部分。
     兩種情況不印小計：
       (a) 該分組一列都沒印出（整組支付價皆為 0 且無不良品註記）→ 整個分組不出現
       (b) 分組代碼留白 → 那是一群「各自獨立、未編分組」的項目，不是一個分組 */
  function subtotalRow(gcode){
    var g = model.groups[gcode];
    var row = new Array(NCOL);
    for (var k=0;k<NCOL;k++) row[k] = xcell('', XS.SUB);
    row[0]  = xcell(gcode, XS.SUB);
    row[6]  = xcell('分組小計', XS.SUB);
    row[10] = xcell(g ? g.name : '', XS.SUB);
    row[14] = xcell(Math.round(gQty[0]), XS.SUBINT);
    row[15] = xcell(Math.round(gAmt[0]), XS.SUBINT);
    row[16] = xcell(Math.round(gQty[1]), XS.SUBINT);
    row[17] = xcell(Math.round(gAmt[1]), XS.SUBINT);
    row[18] = xcell(Math.round(gQty[2]), XS.SUBINT);
    row[19] = xcell(Math.round(gAmt[2]), XS.SUBINT);
    return row;
  }
  function flush(){
    if (!buf.length) return;                       /* (a) 一列都沒印出 → 連小計都不出現 */
    for (var k=0;k<buf.length;k++){
      buf[k][1] = xcell(++seq, buf[k][1] && buf[k][1].s !== undefined ? buf[k][1].s : XS.BODY);
      aoa.push(buf[k]);
    }
    if (cur) aoa.push(subtotalRow(cur));            /* (b) 分組代碼留白 → 不印小計 */
    buf = [];
  }

  var shown = 0;
  for (i=0;i<sel.length;i++){
    var r = sel[i], d = r.d;
    if (cur !== null && r.grp !== cur){ flush(); gAmt = [0,0,0]; gQty = [0,0,0]; }
    cur = r.grp;
    /* 小計涵蓋範圍內全部項目，不受列出規則影響 */
    for (j=0;j<3;j++){ gQty[j] += r.qtyAdj[j]; gAmt[j] += r.amtAdj[j]; }
    if (!isCounted(r)) continue;                   // 不列出，但已計入小計
    shown++;
    var vals = [
      r.grp, 0, atcOf(d), txt(d['ATC名稱']), r.code,
      txt(d['中文名稱']), txt(d['藥品名稱']),
      (r.price === 0 ? '－' : r.price),             /* 暫停支付者不顯示 0，同 Word 表一／表二 */
      txt(d['藥品分類_名稱']), essentialLabel(d['必要藥品']),   /* [D17] 代碼 → 意義 */
      txt(d['分組名稱']), txt(d['劑型']), txt(d['適應症']), txt(d['藥商名稱']),
      Math.round(r.qtyAdj[0]), Math.round(r.amtAdj[0]),
      Math.round(r.qtyAdj[1]), Math.round(r.amtAdj[1]),
      Math.round(r.qtyAdj[2]), Math.round(r.amtAdj[2])
    ];
    buf.push(dataRow(vals, ST, caseCodes[r.code] ? {6: XS.MARK} : null));
  }
  flush();

  var last = xlColName(NCOL-1);
  var rh = {0:26, 2:34};
  if (fixedRow) for (i=3;i<aoa.length;i++) rh[i] = 20;   /* 鎖定列高，長文字只顯示第一行 */
  return {
    name: name, aoa: aoa,
    cols: [11, 6, 10, 20, 13, 18, 28, 10, 13, 9, 30, 12, 38, 18, 11, 13, 11, 13, 11, 13],
    freeze: 3, rowHeights: rh, rowHeight: 20,
    merges: ['B1:' + last + '1', 'B2:' + last + '2'],
    landscape: true, fitToPage: true,
    printArea: '$B$1:$' + last + '$' + aoa.length,   /* 不含 A 欄分組代碼 */
    printTitles: '$1:$3',
    __rows: shown, __scope: sel.length
  };
}

function atcRowsByLen(model, items, len){
  var pres = {}, list = [], i, j;
  for (i=0;i<items.length;i++){
    var a7 = atcOf(items[i].drug);
    if (!a7) continue;
    var pre = a7.substring(0, len);
    if (!pres[pre]){ pres[pre] = 1; list.push(pre); }
  }
  if (!list.length) return [];
  return model.rows.filter(function(r){
    var a = atcOf(r.d);
    if (!a) return false;                                  // 無 ATC7 碼者不納入
    for (var k=0;k<list.length;k++){
      if (len >= 7 ? (a === list[k]) : (a.substring(0, len) === list[k])) return true;
    }
    return false;
  });
}

/* 附件固定產出的 ATC 層級 */
var ATT_ATC_LEVELS = [
  {len:1, tag:'ATC前1碼', label:'相同ATC 前1碼', fixedRow:true},   /* 列數以千計，鎖列高 */
  {len:3, tag:'ATC前3碼', label:'相同ATC 前3碼', fixedRow:true},
  {len:4, tag:'ATC前4碼', label:'相同ATC 前4碼'},
  {len:5, tag:'ATC前5碼', label:'相同ATC 前5碼'},
  {len:7, tag:'ATC7碼',   label:'相同ATC 7碼'}
];

/* levels：要產出的 ATC 碼數陣列，例 [4,5,7]；未指定時全出。「同分組」一律產出。 */
function buildAttachmentWorkbook(model, rep, priceYear, levels){
  var sheets = [], i, n = 1;
  /* 同分組分頁：本案項目所屬分組的「整組」項目。空白分組代碼不得作為擴張依據。 */
  var gset = {};
  for (i=0;i<rep.grpUnion.length;i++) if (rep.grpUnion[i]) gset[rep.grpUnion[i]] = 1;
  var hasBlank = rep.blankGrp > 0;              /* 有本案藥品的分組代碼留白 */
  var grpRows = model.rows.filter(function(r){
    return (r.grp && gset[r.grp]) || (hasBlank && !r.grp && rep.caseCodes[r.code]);
  });
  sheets.push(attachSheet('1.同分組', '查健保收載同分組藥品清單',
    grpRows, model, priceYear, rep.caseCodes));
  for (i=0;i<ATT_ATC_LEVELS.length;i++){
    var lv = ATT_ATC_LEVELS[i];
    if (levels && levels.indexOf(lv.len) < 0) continue;
    sheets.push(attachSheet((++n) + '.' + lv.tag, '查健保收載' + lv.label + '藥品清單'
      + (lv.fixedRow ? '（列高固定，適應症等長欄位僅顯示第一行）' : ''),
      atcRowsByLen(model, rep.items, lv.len), model, priceYear, rep.caseCodes, lv.fixedRow));
  }
  return {book: buildXlsxBook(sheets),
          counts: sheets.map(function(s){ return s.name + ' ' + s.__rows + ' 項'; })};
}
