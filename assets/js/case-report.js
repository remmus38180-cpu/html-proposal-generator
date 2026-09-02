'use strict';
/* 案件需求檔 → 提案報表模型（僅 Path B 載入） */

/* 說明一慣例：不分項，「、」分隔，最後一筆用「及」；同廠商來文集中，只在其第一筆前冠藥商名稱 */
function joinLetters(letters){
  var byV = {}, order = [], i, k, m;
  for (i=0;i<letters.length;i++){
    var v = letters[i].vendor || '';
    if (!byV[v]){ byV[v] = []; order.push(v); }
    byV[v].push(letters[i].ref);
  }
  var parts = [];
  for (k=0;k<order.length;k++){
    var refs = byV[order[k]];
    for (m=0;m<refs.length;m++)
      parts.push((m === 0 ? order[k] : '') + (refs[m] || (BLANK + '年＿月＿日' + BLANK + '字第' + BLANK + '號')) + '函');
  }
  if (!parts.length) return '';
  if (parts.length === 1) return parts[0];
  return parts.slice(0, -1).join('、') + '及' + parts[parts.length-1];
}

function mergeCaseRowsByCode(caseRows){
  var order = [], byCode = {}, i, k;
  for (i=0;i<caseRows.length;i++){
    var cr = caseRows[i], code = txt(cr['CODE']);
    if (!code) continue;
    if (!byCode[code]){ byCode[code] = {CODE:code, __refs:[], __modes:[], __rows:0}; order.push(code); }
    var m = byCode[code];
    m.__rows++;
    for (k in cr){
      if (k === 'CODE' || k === '來文日期文號' || k === 'ATC_MODE') continue;
      if (txt(m[k]) === '' && txt(cr[k]) !== '') m[k] = cr[k];
    }
    var ref = txt(cr['來文日期文號']);
    if (ref && m.__refs.indexOf(ref) < 0) m.__refs.push(ref);
    var md = txt(cr['ATC_MODE']).toUpperCase();
    if (md && m.__modes.indexOf(md) < 0) m.__modes.push(md);
  }
  return order.map(function(c){
    var m = byCode[c], md = m.__modes;
    m['ATC_MODE'] = (md.indexOf('BOTH') >= 0 || (md.indexOf('FULL') >= 0 && md.indexOf('PREFIX5') >= 0))
      ? 'BOTH' : (md[0] || 'BOTH');            /* 未填 → 預設 BOTH */
    if (!m.__refs.length) m.__refs = [''];
    return m;
  });
}

function firstNonEmpty(rows, key){
  for (var i=0;i<rows.length;i++){ var v = txt(rows[i][key]); if (v) return v; }
  return '';
}

/* 「項次」儲存格：本案藥品在數字下方加粗體標記 */
function selfMark(n, isSelf){
  return isSelf ? segCell([{t:String(n)}, {t:'\n（本案藥品）', b:true, sz:16}]) : String(n);
}

function buildCaseReport(caseRows, model, priceYear){
  var priceCol = model.priceCol, YR3 = model.YR3;
  var i, j, k;
  if (!caseRows || !caseRows.length) throw new Error('案件需求檔沒有資料列');

  /* ---- 逐列 → 藥品項目（同 CODE 的多列先合併） ---- */
  var mergedRows = mergeCaseRowsByCode(caseRows);
  var items = [], missing = [];
  for (i=0;i<mergedRows.length;i++){
    var cr = mergedRows[i], code = txt(cr['CODE']);
    if (!code) continue;
    var row = model.byCode[code];
    if (!row){ missing.push(code); continue; }
    var lic = txt(cr['領有許可證'], txt(cr['Has_License'],'')).toUpperCase();
    if (lic !== 'Y' && lic !== 'N') lic = 'Y';
    var ten = num(cr['十國中位價']);
    if (ten === 0) ten = null;                       // 查不到或填 0 → 十國藥價法呈現「無」
    var tenPcts  = parsePcts(cr['十國加計比例']);
    var costPcts = parsePcts(cr['成本加計比例']);
    if (!tenPcts.length)  tenPcts  = [0];
    if (!costPcts.length) costPcts = [0];
    items.push({
      code: code, row: row, drug: row.d, caseRow: cr,
      grp: model.groups[row.grp],
      atcMode: (function(m){ return (m === 'PREFIX5' || m === 'FULL') ? m : 'BOTH'; })(txt(cr['ATC_MODE'],'BOTH').toUpperCase()),
      ten: ten, sug: num(cr['廠商建議價']), cost: num(cr['廠商成本']), lic: lic,
      tenPcts: tenPcts, costPcts: costPcts,
      vendor: txt(row.d['藥商名稱']) || txt(row.d['藥商簡稱']),
      letterRefs: cr.__refs && cr.__refs.length ? cr.__refs : ['']
    });
  }
  if (!items.length) throw new Error('主檔中找不到任何案件 CODE：' + (missing.join('、') || '（需求檔無 CODE 欄）'));

  /* ---- 核價試算 ---- */
  for (i=0;i<items.length;i++){
    var it = items[i];
    it.calc = buildItemCalc(it);
  }

  /* ---- 表一：全案分組代碼聯集，排序 分組代碼↑ → AMT↓ ---- */
  /* ⚠️ 空白分組代碼不得作為「同分組」的擴張依據 —— 主檔若有多筆分組代碼留白，
     它們會被彙總成同一個「空白分組」，一旦擴張就會把大量不相干的藥品拉進來。 */
  var gset = {}, gorder = [], blankGrp = 0;
  for (i=0;i<items.length;i++){
    var gc = items[i].row.grp;
    if (!gc){ blankGrp++; continue; }
    if (!gset[gc]){ gset[gc]=1; gorder.push(gc); }
  }
  var caseCodes = {};
  for (i=0;i<items.length;i++) caseCodes[items[i].code] = 1;

  /* 分組代碼留白的本案藥品，只列自己，不擴張 */
  var t1rows = model.rows.filter(function(r){
    return (r.grp && gset[r.grp]) || (!r.grp && caseCodes[r.code]);
  });
  t1rows.sort(function(a,b){
    if (a.grp !== b.grp) return a.grp < b.grp ? -1 : 1;
    return b.amtAdj[2] - a.amtAdj[2];
  });
  var grpTotal = {};                                  // 占率分母：含 0 元項目
  for (i=0;i<t1rows.length;i++) grpTotal[t1rows[i].grp] = (grpTotal[t1rows[i].grp] || 0) + t1rows[i].amtAdj[2];

  var table1 = [], n1 = 0;
  for (i=0;i<t1rows.length;i++){
    var r = t1rows[i];
    if (!isCounted(r)) continue;                      // 列出規則同 §3.3（統計分母仍為整組）
    var amt = r.amtAdj[2];
    var qty = r.qtyAdj[2];
    var tot = grpTotal[r.grp] || 0;
    /* 欄序：項次／藥品名稱／藥商／現行支付價／申報金額／占率／申報量
       （本案藥品）標記改置於「項次」下方，且僅此四字加粗 */
    var a = [
      selfMark(++n1, caseCodes[r.code]),
      txt(r.d['藥品名稱']),
      txt(r.d['藥商名稱']) || txt(r.d['藥商簡稱']),
      (r.price === null || r.price === 0) ? '－' : money(r.price),   // 暫停支付者不顯示 0
      amt ? fmt(amt) : '－',
      tot ? adaptivePct(amt/tot*100) : '－',
      qty ? fmt(qty) : '－'
    ];
    a.__self = !!caseCodes[r.code];
    table1.push(a);
  }
  /* 表一備註：列出但支付價欄留「－」的暫停支付項目 */
  var t1Susp = t1rows.filter(function(r){ return r.price === 0 && isCounted(r); }).length;
  var table1Notes = t1Susp
    ? ['註：本表含 ' + t1Susp + ' 項暫時停止支付項目，其現行健保支付價欄以「－」表示。']
    : [];

  /* ---- 表二：由本案 ATC 找出分組，再以「整個分組」呈現 ---- */
  /* ATC_MODE = FULL / PREFIX5 / BOTH；BOTH 產生兩張表（7碼全等 + 前5碼相同） */
  function buildT2Block(mode, memberItems, title){
    var g2 = {}, g2order = [], ii, jj;
    for (ii=0;ii<memberItems.length;ii++){
      var atc7 = atcOf(memberItems[ii].drug);
      if (!atc7) continue;
      for (jj=0;jj<model.rows.length;jj++){
        var rr = model.rows[jj];
        if (!rr.grp) continue;                    // 分組代碼留白者不作為擴張依據
        if (!atcMatch(rr.d, atc7, mode)) continue;
        if (!g2[rr.grp]){ g2[rr.grp] = true; g2order.push(rr.grp); }
      }
    }
    var groups = [], anyMultiAtc = false;
    for (ii=0;ii<g2order.length;ii++){
      var gk = g2order[ii];
      var sub = model.rows.filter(function(x){ return x.grp === gk; });     // 整個分組，不論 ATC
      var allZero = sub.every(function(x){ return x.price === 0 || x.price === null; });
      if (allZero) continue;                                                // 全組支付價皆 0 → 不列
      var atcs = {}, atcList = [];
      for (jj=0;jj<sub.length;jj++){ var a7 = atcOf(sub[jj].d); if (!atcs[a7]){ atcs[a7]=1; atcList.push(a7); } }
      atcList.sort();
      if (atcList.length > 1) anyMultiAtc = true;
      var ds0 = sub.map(function(x){ return x.d; });
      groups.push({
        grp: gk, name: txt(sub[0].d['分組名稱']), atcList: atcList, sub: sub,
        self: sub.some(function(x){ return caseCodes[x.code]; }),
        susp: countSuspended(ds0, priceCol)     // 其中屬暫時停止支付者
      });
    }
    groups.sort(function(a,b){ return a.grp < b.grp ? -1 : (a.grp > b.grp ? 1 : 0); });
    /* 分組代碼留白者不成一個分組，表二是分組層的表，故不列入；
       這些項目仍會依 ATC 出現在提案附件的 ATC 分頁（項目層清單）。 */

    /* 同ATC5碼的表格固定顯示 ATC7碼／ATC名稱；同ATC7碼的表格僅在分組跨多碼時顯示 */
    var showAtc = (mode === 'PREFIX5') || anyMultiAtc;

    var head = showAtc
      ? ['項次','分組名稱','ATC7碼','ATC名稱','項目數','現行健保支付價(' + pyLabel(priceYear) + ')(元)',
         YR3+'年整組申報金額(元)', YR3+'年整組申報量']
      : ['項次','分組名稱','項目數','現行健保支付價(' + pyLabel(priceYear) + ')(元)',
         YR3+'年整組申報金額(元)', YR3+'年整組申報量'];
    var widths = showAtc ? [12,21,12,14,7,15,16,13] : [11,28,12,17,17,15];

    var rows = [], nItems = 0;
    for (ii=0;ii<groups.length;ii++){
      var g = groups[ii];
      /* 同分組多種 ATC7 碼 → 分列呈現各自數據；分組名稱與項次以合併儲存格呈現 */
      var buckets = showAtc
        ? g.atcList.map(function(a){
            return {atc:a, rows:g.sub.filter(function(x){ return atcOf(x.d) === a; })};
          })
        : [{atc:null, rows:g.sub}];
      for (jj=0;jj<buckets.length;jj++){
        var bk = buckets[jj], ds = bk.rows.map(function(x){ return x.d; });
        var cnt = countValid(ds, priceCol);        // 價格≠0，或 價格=0 且 不良品暫停支付註記=Y
        nItems += cnt;
        var amt = 0, qty = 0;
        for (var kk=0;kk<bk.rows.length;kk++){ amt += bk.rows[kk].amtAdj[2]; qty += bk.rows[kk].qtyAdj[2]; }
        var tail = [
          String(cnt),
          priceRange(ds, priceCol, '、', true) || '－',            // 支付價不顯示 0 元
          amt ? fmt(Math.round(amt)) : '－',
          qty ? fmt(Math.round(qty)) : '－'
        ];
        var a = showAtc
          ? [selfMark(ii+1, g.self), g.name, bk.atc || '－', txt(ds[0]['ATC名稱']) || '－'].concat(tail)
          : [selfMark(ii+1, g.self), g.name].concat(tail);
        a.__self  = g.self;
        a.__first = (jj === 0);
        a.__span  = buckets.length;
        rows.push(a);
      }
    }
    var notes = [];
    for (ii=0;ii<groups.length;ii++)
      if (groups[ii].susp > 0)
        notes.push('註：' + groups[ii].name + '之項目數包含 ' + groups[ii].susp + ' 項暫時停止支付項目。');
    return {mode:mode, title:title, head:head, widths:widths, rows:rows, notes:notes,
            showAtc:showAtc, nGroups:groups.length, nItems:nItems,
            grpCodes:groups.map(function(g){ return g.grp; })};
  }

  var t2blocks = [];
  var t2full = items.filter(function(x){ return x.atcMode === 'FULL'    || x.atcMode === 'BOTH'; });
  var t2pre5 = items.filter(function(x){ return x.atcMode === 'PREFIX5' || x.atcMode === 'BOTH'; });
  if (t2full.length) t2blocks.push(buildT2Block('FULL',    t2full, '同ATC7碼'));
  if (t2pre5.length) t2blocks.push(buildT2Block('PREFIX5', t2pre5, '同ATC前5碼'));

  /* ---- 表五：廠商建議（每項藥品一列） ---- */
  var table5 = items.map(function(it){
    var tot = (it.row.grp && it.grp) ? it.grp.amtSum[2] : 0;   /* 分組代碼留白 → 占率印「－」 */
    return [
      txt(it.drug['藥品名稱']),
      txt(it.drug['分組名稱']),
      it.row.price === null ? '－' : money(it.row.price),
      it.sug === null ? BLANK : money(it.sug),
      tot ? adaptivePct(it.row.amtAdj[2]/tot*100) : '－',
      it.cost === null ? BLANK : money(it.cost),
      it.ten === null ? '無' : money(it.ten)
    ];
  });

  /* ---- 來文（藥商 + 文號）去重 ---- */
  var seenL = {}, letters = [];
  for (i=0;i<items.length;i++){
    var refs = items[i].letterRefs;
    for (j=0;j<refs.length;j++){
      var key = items[i].vendor + '|' + refs[j];
      if (seenL[key]) continue;
      seenL[key] = 1;
      letters.push({vendor: items[i].vendor, ref: refs[j]});
    }
  }

  /* ---- 第40條：全案彙總 ---- */
  var lt5 = 0, l550 = 0, ge50 = 0;
  for (i=0;i<items.length;i++){
    var vs = items[i].calc.allPrices;
    for (j=0;j<vs.length;j++){
      var v = vs[j];
      if (v === null) continue;
      if (v < 5) lt5 = 1; else if (v < 50) l550 = 1; else ge50 = 1;
    }
  }

  /* ---- 註1／註2：相同者只印一次 ---- */
  var noteMap = {}, notes = [];
  for (i=0;i<items.length;i++){
    var c = items[i].calc;
    var nk = c.a10Law + '|' + c.costLaw;
    if (!noteMap[nk]){ noteMap[nk] = {a10Law:c.a10Law, costLaw:c.costLaw, idx:[], lics:{}, licOrder:[]}; notes.push(noteMap[nk]); }
    var nn = noteMap[nk];
    nn.idx.push(i+1);
    if (!nn.lics[c.licNote]){ nn.lics[c.licNote] = []; nn.licOrder.push(c.licNote); }
    nn.lics[c.licNote].push(i+1);
  }

  return {
    items: items, missing: missing, priceYear: priceYear, priceCol: priceCol,
    baseYear: model.baseYear, dataYear: YR3, model: model,
    reporter: firstNonEmpty(caseRows, '報告人') || BLANK,
    letters: letters, grpUnion: gorder, caseCodes: caseCodes, blankGrp: blankGrp,
    table1: table1, table1Notes: table1Notes, table2Blocks: t2blocks, table5: table5,
    table1Head: T1H(YR3, priceYear), table1W: T1W,
    table5Head: T5H(YR3, priceYear), table5W: T5W,
    letterSentence: joinLetters(letters),
    notes: notes, art40: article40Text(lt5, l550, ge50),
    subject: '有關' + items[0].vendor + '藥品「'
           + txt(items[0].drug['藥品名稱']) + '」建議調高其健保支付價格案，提請討論。'
  };
}

/* 單一藥品的核價試算矩陣 */
function buildItemCalc(it){
  /* 分組代碼留白的藥品沒有可用的分組層統計 —— 不可拿「空白分組」的彙總當級距與平均申報量 */
  var grp = (it.row.grp && it.grp) ? it.grp : {};
  var avg = (grp.avgQty3 === undefined) ? null : grp.avgQty3;
  var nhi = it.row.price;
  var tenVals  = (it.ten === null) ? [null] : it.tenPcts.map(function(p){ return a10RefPrice(it.ten, p/100); });
  var tenHead  = (it.ten === null) ? ['－']  : it.tenPcts.map(function(p){ return p + '%'; });
  var costVals = (it.cost === null) ? it.costPcts.map(function(){ return null; })
                                    : it.costPcts.map(function(p){ return costRefPrice(it.cost, p/100, it.lic); });
  var costHead = it.costPcts.map(function(p){ return p + '%'; });

  function pTxt(v, noTen){ return noTen ? '無' : (v === null ? '－' : String(v)); }
  function amtTxt(v, noTen){
    if (noTen) return '無';
    if (v === null || avg === null) return '－';
    return formatAmt(v * avg);
  }
  function impTxt(v, noTen){
    if (noTen) return '無';
    if (v === null || avg === null || nhi === null) return '－';
    return formatAmt((v - nhi) * avg);
  }
  var noTen = (it.ten === null);
  var tier = amtTier(grp.monthlyAmt === undefined ? null : grp.monthlyAmt);
  return {
    tenHead: tenHead, costHead: costHead,
    tenVals: tenVals, costVals: costVals,
    priceRow:  tenVals.map(function(v){ return pTxt(v, noTen); }).concat(costVals.map(function(v){ return pTxt(v); })),
    totalRow:  tenVals.map(function(v){ return amtTxt(v, noTen); }).concat(costVals.map(function(v){ return amtTxt(v); })),
    impactRow: tenVals.map(function(v){ return impTxt(v, noTen); }).concat(costVals.map(function(v){ return impTxt(v); })),
    allPrices: (noTen ? [] : tenVals).concat(costVals),
    noTen: noTen, avgQty: avg, monthlyAmt: grp.monthlyAmt,
    tier: tier,
    a10Pct:  tier === null ? null : A10_TIER_PCT[tier],
    costPct: tier === null ? null : COST_TIER_PCT[tier],
    a10Law:  tier === null ? '' : '參考十國藥價法' + A10_TIER_LAW[tier],
    costLaw: tier === null ? '' : '參考成本價法'   + COST_TIER_LAW[tier],
    licNote: LIC_NOTE[String(it.lic).toUpperCase()] || ''
  };
}

/* 表一：刪除「整組每月平均申報金額」，申報量移至最後；金額／數量欄名帶年度 */
function T1H(yr3, py){
  return ['項次','藥品名稱','藥商','現行健保支付價(' + pyLabel(py) + ')(元)',
          yr3 + '年整組申報金額(元)', yr3 + '年占率', yr3 + '年整組申報量'];
}

var T1W = [11,21,12,14,14,8,12];

function T5H(yr3, py){
  return ['藥品名稱','分類分組名稱','現行支付價(' + pyLabel(py) + ')(元)','廠商建議價(元)',
          yr3 + '年廠商占率','廠商成本(元)','十國藥價中位數(元)'];
}

var T5W = [21,23,13,12,10,10,12];

/* 核價試算小表（合併儲存格） */
function calcTableRows(it, priceYear){
  var c = it.calc;
  var n1 = c.tenHead.length, n2 = c.costHead.length, i;
  var hd = {shade:SHADE_HD, b:true};
  var r1 = [{t:'健保價(' + pyLabel(priceYear) + ')', span:1, shade:SHADE_HD, b:true},
            {t:'廠商建議價', span:1, shade:SHADE_HD, b:true},
            {t:'核價方式', span:n1+n2, shade:SHADE_HD, b:true}];
  /* 註1／註2 以上標呈現 */
  var r2 = [{t: it.row.price === null ? '－' : money(it.row.price)},
            {t: it.sug === null ? '－' : money(it.sug)},
            {t: segCell([{t:'十國藥價中位數', b:true}, {t:'註1', b:true, sup:true}]),
             span:n1, shade:SHADE_HD, b:true},
            {t: segCell([{t:'參考成本價', b:true}, {t:'註2', b:true, sup:true}]),
             span:n2, shade:SHADE_HD, b:true}];
  var r3 = [{t:'加計比例', span:2, shade:SHADE_HD, b:true}];
  for (i=0;i<n1;i++) r3.push({t:c.tenHead[i], shade:SHADE_HD, b:true});
  for (i=0;i<n2;i++) r3.push({t:c.costHead[i], shade:SHADE_HD, b:true});
  function dataRow(label, vals){
    var r = [{t:label, span:2, shade:SHADE_HD, b:true}];
    for (var k=0;k<vals.length;k++) r.push({t:vals[k]});
    return r;
  }
  var weights = [13,13];
  for (i=0;i<n1+n2;i++) weights.push(74/(n1+n2));
  return {rows:[r1, r2, r3,
                dataRow('調高後支付價', c.priceRow),
                dataRow('整體藥費',     c.totalRow),
                dataRow('財務衝擊',     c.impactRow)],
          weights: weights};
}
