'use strict';
/* 提案 Word 組字（僅 Path B 載入） */

var L0 = {numId:NUM_ID, ilvl:0};

var L1 = {numId:NUM_ID, ilvl:1};

var L2 = {numId:NUM_ID, ilvl:2};

function lv(base, extra){
  var o = {}, k;
  for (k in base) o[k] = base[k];
  if (extra) for (k in extra) o[k] = extra[k];
  return o;
}

var GAP = function(){ return para('', {sz:16, line:160}); };

/* 表二 → docTableRaw 用的儲存格陣列；同分組跨多 ATC7 碼時，項次與分組名稱垂直合併 */
function t2RawRows(bk){
  var out = [bk.head.map(function(h){ return {t:h, shade:SHADE_HD, b:true}; })];
  for (var r=0;r<bk.rows.length;r++){
    var src = bk.rows[r], multi = (src.__span || 1) > 1;
    var al = bk.showAtc
      ? ['center','left','center','left','center','right','right','right']
      : ['center','left','center','right','right','right'];
    var cells = src.map(function(v, i){
      var c = {t:v, align: al[i] || 'center'};
      if (multi && i <= 1) c.vm = src.__first ? 'restart' : 'cont';
      return c;
    });
    out.push(cells);
  }
  return out;
}

function reportToDocx(rep){
  var b = '', i, j;
  var YR3 = rep.dataYear;

  b += paraTab('提案', '報告人：' + rep.reporter, {b:true, after:120});
  b += para('案由：' + rep.subject, {indent:640, hanging:640, after:120});
  b += para('說明：', {after:60});

  /* 一、依據 — 不分項，「、」分隔，末筆以「及」連接 */
  b += para('依據' + rep.letterSentence + '辦理。', L0);

  /* 二、藥品資訊 + 表1 表2 */
  b += para('本案藥品健保收載情形說明如下：', L0);
  for (i=0;i<rep.items.length;i++){
    var it = rep.items[i], d = it.drug;
    var cls = drugClassLabel(d);
    var ord = catOrdinal(d['分類']);
    var yrs = (function(){ var ly = ad8(d['收載日期']); return ly === null ? null : rep.baseYear - ly; })();
    var listed = yrs === null ? '超過／未滿' : (yrs >= 15 ? '超過' : '未滿');
    var atc7 = atcOf(d), atcNm = txt(d['ATC名稱']);
    b += para(txt(d['藥品名稱']) + '（健保代碼 ' + it.code + '）之主成分、劑型、含量為「'
              + (txt(d['分組名稱']) || BLANK) + '」，屬' + (cls || BLANK)
              + '，屬健保收載' + listed + '15年之' + (ord || '＿') + '大類藥品，適應症「'
              + (txt(it.caseRow['藥品治療用途']) || txt(d['適應症']) || BLANK) + '」，ATC7碼為'
              + (atc7 || '＿') + (atcNm ? ' ' + atcNm : '') + '。', L1);
  }
  b += para('查健保收載同分組藥品共計' + rep.table1.length + '項(P-)：', lv(L1, {after:40}));
  b += docTable(rep.table1Head, rep.table1, rep.table1W);
  for (j=0;j<rep.table1Notes.length;j++) b += para(rep.table1Notes[j], {sz:22, indent:640, after:0});
  b += GAP();
  for (i=0;i<rep.table2Blocks.length;i++){
    var bk = rep.table2Blocks[i];
    var lbl = bk.mode === 'PREFIX5' ? '相同ATC 5碼' : '相同ATC 7碼';
    b += para('查健保收載' + lbl + '藥品共計' + bk.nGroups + '分組、' + bk.nItems + '項目(P-)：',
              lv(L1, {after:40}));
    b += docTableRaw(t2RawRows(bk), bk.widths);
    for (j=0;j<bk.notes.length;j++) b += para(bk.notes[j], {sz:22, indent:640, after:0});
    b += GAP();
  }

  /* 三、四 — 留白讓承辦人填 */
  b += para('經洽詢廠商供應狀況，本藥品' + BLANK + BLANK + '。', L0);
  b += para('經詢問其他廠商，' + BLANK + BLANK + BLANK + '。', L0);

  /* 五、廠商建議 */
  b += para('今' + rep.items[0].vendor + '等來函建議提高健保支付價(P-)：', lv(L0, {after:40}));
  b += docTable(rep.table5Head, rep.table5, rep.table5W);
  b += GAP();

  /* 六、核價試算 */
  b += para('依「全民健康保險藥物給付項目及支付標準」第35條特殊藥品支付價格訂定原則，'
          + '本案藥品就不同核價方式試算之支付價格及財務衝擊如下：（單位：元；財務衝擊為年度推估）', L0);
  for (i=0;i<rep.items.length;i++){
    var itm = rep.items[i];
    b += para(txt(itm.drug['藥品名稱']) + '（藥品代碼：' + itm.code + '）', lv(L1, {after:40}));
    var ct = calcTableRows(itm, rep.priceYear);
    b += docTableRaw(ct.rows, ct.weights);
    /* 表格下方備註：僅保留級距判定用的每月申報金額（申報量已移至 註3/4） */
    var mo = itm.calc.monthlyAmt;
    b += para('　整組每月平均申報金額(' + YR3 + '年)：'
            + (mo === null || mo === undefined ? '－' : fmt(Math.round(mo))) + ' 元',
            {sz:22, indent:960});
    b += GAP();
  }
  /* 註1／註2 */
  for (i=0;i<rep.notes.length;i++){
    var nt = rep.notes[i];
    var tag = rep.notes.length > 1 ? '（項次 ' + nt.idx.join('、') + '）' : '';
    b += para('註1：' + tag + nt.a10Law, {sz:22, indent:1200, hanging:560});
    b += para('註2：' + tag + nt.costLaw, {sz:22, indent:1200, hanging:560});
    b += para('2.領有藥物許可證者，得加計繳納藥害救濟徵收金比率及營業稅。', {sz:22, indent:1200});
    for (j=0;j<nt.licOrder.length;j++){
      var lt = nt.licOrder.length > 1 ? '（項次 ' + nt.lics[nt.licOrder[j]].join('、') + '）' : '';
      b += para('  ' + lt + nt.licOrder[j], {sz:22, indent:1200});
    }
  }
  b += para(rep.art40, {sz:22, indent:960, hanging:0});
  /* 註3／註4：含每個項目的實際申報量數字 */
  for (i=0;i<rep.items.length;i++){
    var itm3 = rep.items[i];
    var av3 = itm3.calc.avgQty;
    var avStr3 = (av3 === null || av3 === undefined) ? '－' : fmt(Math.round(av3));
    var tag3 = rep.items.length > 1 ? '（項次 ' + (i+1) + '）' : '';
    b += para('註3：' + tag3 + '整體藥費＝調高後支付價 × 同分組近三年平均申報量(QTY) ＝ 調高後支付價 × ' + avStr3,
              {sz:22, indent:1200, hanging:560});
    b += para('註4：' + tag3 + '財務衝擊＝(調高後支付價 － 健保價) × 同分組近三年平均申報量(QTY) ＝ (調高後支付價 － 健保價) × ' + avStr3,
              {sz:22, indent:1200, hanging:560});
  }
  b += GAP();

  /* 七、八 */
  b += para('本案經諮詢藥品專家之意見如下：', L0);
  b += para('專家意見一(P-)：' + BLANK + BLANK, L1);
  b += para('專家意見二(P-)：' + BLANK + BLANK, L1);
  b += para('專家意見三(P-)：' + BLANK + BLANK, L1);
  b += para('本案是否提高藥價及方式？提請全體諮詢專家共同討論。', L0);

  return buildDocx(b);
}
