'use strict';
/* 主控：DUAL_TRACK 狀態管理、路徑 A / B 分流、模組權限檢查 */

var $ = function(id){ return document.getElementById(id); };

/* ============================================================
   全域狀態
   ============================================================ */
var DUAL_TRACK = {
  path: null,                    // 'PATH_A' | 'PATH_B'
  priceYearMonth: null,          // '11508'

  /* 共用層（Path A/B 都用） */
  masterData: {
    parsed: null,                // xlsx 原始解析結果
    model: null,                 // buildGroupModel 結果，含 [B6]/[B7]/[B8]
    exclusions: { tpn: [], offnet: [] },
    priceCols: []                // 主檔偵測到的 PRICE 年月清單
  },

  /* 路徑 A */
  pathA: {
    sheets: { sheet1: null, sheet2: null, sheet3: null }
  },

  /* 路徑 B */
  pathB: {
    caseData: null,              // 原始案件需求檔列
    targets: null,               // buildCaseReport 結果（含試算）
    outputs: { word: null, xlsxMain: null, xlsxAttach: null }
  }
};

/* 供舊模組沿用的別名（buildCaseReport 等以 model 為參數，不直接讀這裡） */
var MODEL = null;

/* ============================================================
   基礎輸出
   ============================================================ */
function log(msg, kind){
  var box = $('logBox');
  var span = document.createElement('span');
  if (kind) span.className = 'l-'+kind;
  span.textContent = msg + '\n';
  box.appendChild(span);
  box.scrollTop = box.scrollHeight;
}
function setState(el, msg, kind){
  el.textContent = msg;
  el.className = 'file-state' + (kind ? ' '+kind : '');
}

/* 下載：本機開檔用 <a download>；線上沙箱版走 downloads 能力 */
var DL = null;
if (window.claude && typeof window.claude.use === 'function'){
  window.claude.use('downloads').then(function(d){ DL = d; markSaveMode(); })
                                .catch(function(){ markSaveMode(); });
}
function anchorSave(blob, name){
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(function(){ URL.revokeObjectURL(url); }, 4000);
}
function save(blob, name, alt){
  if (blob && typeof blob.then === 'function'){
    return blob.then(function(b){ save(b, name, alt); })
               .catch(function(e){ log('產生失敗：' + name + ' — ' + e.message, 'bad'); });
  }
  if (!DL) return anchorSave(blob, name);
  blob.arrayBuffer().then(function(ab){ return DL.save({filename:name, data:ab}); })
    .catch(function(e){
      var c = e && e.code;
      if ((c === 'rejected_extension' || c === 'extension_not_enabled') && alt){
        log('線上版不允許 .' + name.split('.').pop() + '，改存 ' + alt.name, 'warn');
        save(alt.blob, alt.name);
      } else if (c === 'declined'){
        log('已取消儲存：' + name, 'warn');
      } else {
        log('儲存失敗（' + (c || 'unknown') + '）：' + name + '　→ 請改用本機檔案', 'bad');
      }
    });
}
function toCsv(rows){
  var head = Object.keys(rows[0] || {});
  function cell(v){
    v = (v === undefined || v === null) ? '' : String(v);
    return /[",\n]/.test(v) ? '"' + v.replace(/"/g,'""') + '"' : v;
  }
  var out = [head.join(',')];
  for (var i=0;i<rows.length;i++) out.push(head.map(function(h){ return cell(rows[i][h]); }).join(','));
  return out.join('\r\n');
}
function csvBlob(rows){ return new Blob(['\uFEFF' + toCsv(rows)], {type:'text/csv;charset=utf-8'}); }

(function(){
  var b = $('badgeEngine');
  if (typeof DecompressionStream !== 'undefined'){
    b.className = 'badge on';
    b.innerHTML = '<span class="dot"></span>xlsx 解析引擎就緒';
  } else {
    b.className = 'badge off';
    b.innerHTML = '<span class="dot"></span>瀏覽器太舊，請改用 CSV';
  }
})();
function markSaveMode(){
  var id = 'badgeSave', old = $(id);
  if (old) old.remove();
  var n = document.createElement('span');
  n.id = id;
  if (DL){ n.className = 'badge on'; n.innerHTML = '<span class="dot"></span>可直接儲存檔案'; }
  else { n.className = 'badge off'; n.innerHTML = '<span class="dot"></span>沙箱擋下載 — 請用本機檔案'; }
  $('badgeEngine').parentNode.appendChild(n);
}
if (window.claude && typeof window.claude.use === 'function') markSaveMode();

/* ============================================================
   路徑 B 模組：延遲載入
   路徑 A 全程不會走到這裡，price-calc / case-report / docx-writer /
   docx-report / attachment-workbook / template 因此不會出現在頁面上。
   ============================================================ */
var PATH_B_MODULES = [
  'assets/js/price-calc.js',
  'assets/js/case-report.js',
  'assets/js/docx-writer.js',
  'assets/js/exporters/docx-report.js',
  'assets/js/exporters/attachment-workbook.js',
  'assets/js/template.js'
];
var _pathBPromise = null;
function pathBLoaded(){ return typeof buildCaseReport === 'function'; }

/* 分檔版：以 <script src> 載入。
   單檔版（build/bundle.mjs 產生）：原始碼已內嵌為 <script type="text/plain" data-mod="…">，
   改成把該區塊的文字丟進新的 <script> 執行。兩種形式的**執行時機完全相同** ——
   路徑 A 全程不會執行到這些模組，pathBLoaded() 在選路徑 B 之前一律為 false。 */
function injectModule(src){
  return new Promise(function(res, rej){
    var emb = document.querySelector('script[type="text/plain"][data-mod="' + src + '"]');
    var s = document.createElement('script');
    if (emb){
      s.textContent = emb.textContent;
      document.body.appendChild(s);        // 同步執行
      res(); return;
    }
    s.src = src; s.async = false;
    s.onload = function(){ res(); };
    s.onerror = function(){ rej(new Error('載入失敗：' + src)); };
    document.body.appendChild(s);
  });
}
function loadPathBModules(){
  if (_pathBPromise) return _pathBPromise;
  _pathBPromise = PATH_B_MODULES.reduce(function(chain, src){
    return chain.then(function(){ return injectModule(src); });
  }, Promise.resolve()).then(function(){
    log('路徑 B 模組已載入（' + PATH_B_MODULES.length + ' 個）。', 'ok');
    bindPathBHandlers();
    refresh();
  }).catch(function(e){
    _pathBPromise = null;
    log(e.message + '　→ 請確認 assets/js 目錄與 index.html 在同一層。', 'bad');
    throw e;
  });
  return _pathBPromise;
}

/* ============================================================
   狀態清空規則（INTENT「清空規則」表）
   ============================================================ */
function clearAll(){
  DUAL_TRACK.path = null;
  DUAL_TRACK.priceYearMonth = null;
  DUAL_TRACK.masterData = {parsed:null, model:null, exclusions:{tpn:[],offnet:[]}, priceCols:[]};
  clearPathA(); clearPathBAll();
  MODEL = null; _modelKey = null;
  if ($('rdA')) $('rdA').checked = false;
  if ($('rdB')) $('rdB').checked = false;
  $('prevBox').innerHTML = '';
}
function clearModelOnly(){                 /* 改選 price 年月：檔案留著，模型與兩路徑產出全清 */
  DUAL_TRACK.masterData.model = null;
  MODEL = null; _modelKey = null;
  clearPathA(); clearPathBOutputs();
  DUAL_TRACK.pathB.targets = null;
  $('prevBox').innerHTML = '';
}
function clearPathA(){ DUAL_TRACK.pathA.sheets = {sheet1:null, sheet2:null, sheet3:null}; }
function clearPathBOutputs(){ DUAL_TRACK.pathB.outputs = {word:null, xlsxMain:null, xlsxAttach:null}; }
function clearPathBAll(){ DUAL_TRACK.pathB.caseData = null; DUAL_TRACK.pathB.targets = null; clearPathBOutputs(); }

/* ============================================================
   主檔驗證與 PRICE 年月掃描
   ============================================================ */
var MASTER_REQUIRED = ['CODE','分組代碼','ATC7碼'];
function scanPriceCols(rows){
  var keys = {}, i, k;
  for (i=0;i<Math.min(rows.length,300);i++) for (k in rows[i]) keys[k]=1;
  var out = [];
  for (k in keys){
    var m = /^PRICE(\d{5})$/.exec(k);
    if (m) out.push(m[1]);
  }
  out.sort();
  return out;
}
/* 回傳 {ok:Boolean, missing:[], priceCols:[]} */
function validateMaster(rows){
  var keys = {}, i, k;
  for (i=0;i<Math.min(rows.length,300);i++) for (k in rows[i]) keys[k]=1;
  var missing = MASTER_REQUIRED.filter(function(c){ return !keys[c]; });
  var priceCols = scanPriceCols(rows);
  if (!priceCols.length) missing.push('PRICE{民國年月}');
  return {ok: missing.length === 0, missing: missing, priceCols: priceCols};
}

function currentPy(){ return normPriceYear($('priceYear').value); }
function priceColExists(py){
  var m = DUAL_TRACK.masterData.parsed;
  if (!m || py.length !== 5) return false;
  return DUAL_TRACK.masterData.priceCols.indexOf(py) >= 0
      || m.some(function(r){ return ('PRICE'+py) in r; });
}
/* price 年月無對應欄時，給最接近的兩個候選 */
function nearestPriceCols(py){
  var list = DUAL_TRACK.masterData.priceCols;
  if (!list.length) return [];
  var n = parseInt(py,10);
  return list.slice().sort(function(a,b){
    return Math.abs(parseInt(a,10)-n) - Math.abs(parseInt(b,10)-n);
  }).slice(0,3);
}

/* ============================================================
   分組模型（[A3] 排除 → buildGroupModel）
   ============================================================ */
var _modelKey = null, _excl = null;
function caseCodeSet(){
  var m = {}, cs = DUAL_TRACK.pathB.caseData;
  if (cs) for (var i=0;i<cs.length;i++){ var c = txt(cs[i]['CODE']); if (c) m[c] = 1; }
  return m;
}
function getModel(){
  var py = currentPy(), master = DUAL_TRACK.masterData.parsed;
  if (!master || py.length !== 5 || !priceColExists(py)) return null;
  var exTpn = !$('exTpn') || $('exTpn').checked;
  var exOff = !$('exOff') || $('exOff').checked;
  /* 路徑 A 沒有案件檔，keepCodes 為空 —— 這正是兩條路徑母體可能不同的唯一來源 */
  var keep = (DUAL_TRACK.path === 'PATH_B') ? caseCodeSet() : {};
  var key = py + '|' + master.length + '|' + exTpn + exOff + '|' + DUAL_TRACK.path + '|' + Object.keys(keep).join(',');
  if (MODEL && _modelKey === key) return MODEL;
  try {
    var ex = excludeMaster(master, exTpn, exOff, keep);
    _excl = ex;
    DUAL_TRACK.masterData.exclusions = {tpn: ex.tpn, offnet: ex.offnet};
    var baseYear = parseInt(py.substring(0,3),10);
    MODEL = buildGroupModel(ex.rows, py, resolveYearWindow(ex.rows, baseYear));
    DUAL_TRACK.masterData.model = MODEL;
    DUAL_TRACK.priceYearMonth = py;
    _modelKey = key;
    return MODEL;
  } catch (err){
    MODEL = null; _modelKey = null;
    log('分組模型建立失敗：' + err.message + '　→ 請檢視主檔原始資料（診斷資訊如上）。', 'bad');
    return null;
  }
}

/* ============================================================
   檔案載入
   ============================================================ */
function loadFile(file){
  if (/\.csv$/i.test(file.name)) return file.text().then(readCsv);
  return file.arrayBuffer().then(readXlsx);
}

$('fMaster').addEventListener('change', function(e){
  var file = e.target.files && e.target.files[0];
  e.target.value = '';
  if (!file) return;
  clearAll(); refresh();
  setState($('sMaster'), '讀取中…');
  loadFile(file).then(function(rows){
    if (!rows.length) throw new Error('檔案沒有資料列');
    var v = validateMaster(rows);
    if (!v.ok){
      setState($('sMaster'), '✗ 缺少必要欄位', 'bad');
      log('藥品主檔缺少必要欄位：' + v.missing.join('、') + '　→ 請重新上傳。', 'bad');
      refresh();
      return;
    }
    DUAL_TRACK.masterData.parsed = rows;
    DUAL_TRACK.masterData.priceCols = v.priceCols;
    setState($('sMaster'), '✓ ' + rows.length.toLocaleString() + ' 筆 · ' + Object.keys(rows[0]).length + ' 欄', 'ok');
    log('藥品主檔載入成功：' + file.name + '（' + rows.length + ' 筆）', 'ok');
    log('偵測到 ' + v.priceCols.length + ' 個 PRICE 年月欄位：'
        + v.priceCols.join('、') + '　→ 請於步驟 02 選擇。', 'ok');
    fillPySelect(v.priceCols);
    warnMasterQuality(rows);
    refresh();
  }).catch(function(err){
    setState($('sMaster'), '✗ 讀取失敗', 'bad');
    log('藥品主檔讀取失敗：' + err.message, 'bad');
    refresh();
  });
});

function fillPySelect(cols){
  var sel = $('pySelect');
  sel.innerHTML = '';
  var o0 = document.createElement('option');
  o0.value = ''; o0.textContent = cols.length ? '— 請選擇 —' : '— 主檔無 PRICE 欄位 —';
  sel.appendChild(o0);
  cols.forEach(function(c){
    var o = document.createElement('option');
    o.value = c; o.textContent = c + '（' + pyLabel(c) + '）';
    sel.appendChild(o);
  });
  /* 只有一個年月時直接選起來 */
  if (cols.length === 1){ sel.value = cols[0]; $('priceYear').value = cols[0]; }
}

function warnMasterQuality(rows){
  var nBlankG = 0, nBlankA = 0;
  for (var i=0;i<rows.length;i++){
    if (!txt(rows[i]['分組代碼'])) nBlankG++;
    if (!txt(rows[i]['ATC7碼'])) nBlankA++;
  }
  if (nBlankG) log('主檔有 ' + nBlankG.toLocaleString() + ' 筆「分組代碼」留白，'
    + '不會被當成同一分組擴張，[B6]／[B8] 對這些列無效。', 'warn');
  if (nBlankA) log('主檔有 ' + nBlankA.toLocaleString() + ' 筆「ATC7碼」留白，不會出現在任何 ATC 分頁。', 'warn');
}

/* ---- 案件需求檔（路徑 B） ---- */
function bindCaseInput(){
  var el = $('fCase');
  if (!el || el.__bound) return;
  el.__bound = 1;
  el.addEventListener('change', function(e){
    var file = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!file) return;
    DUAL_TRACK.pathB.caseData = null;
    clearPathBOutputs(); DUAL_TRACK.pathB.targets = null;
    MODEL = null; _modelKey = null;
    refresh();
    setState($('sCase'), '讀取中…');
    loadFile(file).then(function(rows){
      if (!rows.length) throw new Error('檔案沒有資料列');
      DUAL_TRACK.pathB.caseData = rows;
      setState($('sCase'), '✓ ' + rows.length.toLocaleString() + ' 筆 · ' + Object.keys(rows[0]).length + ' 欄', 'ok');
      log('案件需求檔載入成功：' + file.name + '（' + rows.length + ' 筆）', 'ok');
      warnCaseMatch(rows);
      MODEL = null; _modelKey = null;
      refresh();
    }).catch(function(err){
      DUAL_TRACK.pathB.caseData = null;
      setState($('sCase'), '✗ 讀取失敗', 'bad');
      log('案件需求檔讀取失敗：' + err.message, 'bad');
      refresh();
    });
  });
}
/* fCase 與 fMaster 無匹配 CODE → 提示（該列稍後由 buildCaseReport 移除） */
function warnCaseMatch(rows){
  var master = DUAL_TRACK.masterData.parsed;
  if (!master) return;
  var have = {};
  for (var i=0;i<master.length;i++) have[txt(master[i]['CODE'])] = 1;
  var miss = [], seen = {};
  for (i=0;i<rows.length;i++){
    var c = txt(rows[i]['CODE']);
    if (!c || seen[c]) continue;
    seen[c] = 1;
    if (!have[c]) miss.push(c);
  }
  if (miss.length) log('這些 CODE 不在主檔中，產出時會被移除：' + miss.join('、'), 'warn');
}

/* ============================================================
   步驟 02 / 03 互動
   ============================================================ */
$('pySelect').addEventListener('change', function(e){
  $('priceYear').value = e.target.value;
  clearModelOnly();
  refresh();
});
$('priceYear').addEventListener('input', function(e){
  e.target.value = normPriceYear(e.target.value).substring(0,5);
  var sel = $('pySelect');
  sel.value = (DUAL_TRACK.masterData.priceCols.indexOf(e.target.value) >= 0) ? e.target.value : '';
  clearModelOnly();
  refresh();
});
['exTpn','exOff'].forEach(function(id){
  if ($(id)) $(id).addEventListener('change', function(){ clearModelOnly(); refresh(); });
});

function setTrack(next){
  var prev = DUAL_TRACK.path;
  if (prev === next) return;
  if (prev === 'PATH_B' && next === 'PATH_A' &&
      (DUAL_TRACK.pathB.caseData || DUAL_TRACK.pathB.targets)){
    if (!window.confirm('切換到路徑 A 會清除已載入的案件需求檔與提案產出，確定嗎？')){
      $('rdB').checked = true;
      return;
    }
  }
  DUAL_TRACK.path = next;
  if (next === 'PATH_A'){
    clearPathBAll();
    if ($('sCase')) setState($('sCase'), '尚未載入');
    $('prevBox').innerHTML = '';
  } else {
    clearPathA();
  }
  MODEL = null; _modelKey = null;
  log('已選擇' + (next === 'PATH_A' ? '路徑 A｜背景查詢' : '路徑 B｜正式提案'), 'ok');
  if (next === 'PATH_B' && !pathBLoaded()) loadPathBModules().catch(function(){});
  refresh();
}
$('rdA').addEventListener('change', function(){ if (this.checked) setTrack('PATH_A'); });
$('rdB').addEventListener('change', function(){ if (this.checked) setTrack('PATH_B'); });

/* ============================================================
   畫面更新
   ============================================================ */
function refresh(){
  var master = DUAL_TRACK.masterData.parsed;
  var py = currentPy();
  var pyOk = !!(master && py.length === 5 && priceColExists(py));

  $('badgeTrack').className = 'badge' + (DUAL_TRACK.path ? ' on' : '');
  $('badgeTrack').innerHTML = '<span class="dot"></span>'
    + (DUAL_TRACK.path === 'PATH_A' ? '路徑 A｜背景查詢'
     : DUAL_TRACK.path === 'PATH_B' ? '路徑 B｜正式提案' : '尚未選擇路徑');

  /* 步驟 02 診斷 */
  var d = $('diagBox');
  if (!master || py.length !== 5){ d.hidden = true; }
  else {
    d.hidden = false;
    var priceCol = 'PRICE' + py;
    $('dPriceCol').textContent = priceCol + (pyOk ? ' ✓' : ' ✗ 主檔無此欄');
    var baseYear = parseInt(py.substring(0,3),10);
    var yr3 = resolveYearWindow(master, baseYear);
    $('dDataYear').textContent = '民國 ' + (yr3-2) + '–' + yr3 + ' 年';
    $('dMasterN').textContent  = master.length.toLocaleString();
    var mm = parseInt(py.substring(3,5),10);
    $('dCutoff').textContent = '民國 ' + ((mm >= 4) ? baseYear : baseYear-1) + ' 年（' + (mm>=4?'≥4月取當年':'<4月取前年') + '）';
    if (pyOk) getModel();
    $('dExcl').textContent = _excl
      ? ((_excl.tpn + _excl.offnet)
          ? ('TPN ' + _excl.tpn.toLocaleString() + ' 筆／不上網 ' + _excl.offnet.toLocaleString()
             + ' 筆　→ 實際採用 ' + (master.length - _excl.tpn - _excl.offnet).toLocaleString() + ' 筆')
          : '無')
      : '—';
    var cs = DUAL_TRACK.pathB.caseData;
    if (!cs) $('dCaseN').textContent = '—';
    else {
      var seen = {}, n = 0;
      for (var ci=0; ci<cs.length; ci++){
        var cc = txt(cs[ci]['CODE']);
        if (cc && !seen[cc]){ seen[cc] = 1; n++; }
      }
      $('dCaseN').textContent = n + ' 項藥品' + (cs.length > n ? '（' + cs.length + ' 列，同 CODE 已合併）' : '');
    }
  }

  /* 步驟 03：price 年月驗證通過才解鎖路徑選擇 */
  $('stepTrack').hidden = !pyOk;
  $('rdA').disabled = !pyOk;
  $('rdB').disabled = !pyOk;
  $('trackA').classList.toggle('sel', DUAL_TRACK.path === 'PATH_A');
  $('trackB').classList.toggle('sel', DUAL_TRACK.path === 'PATH_B');
  $('casePanel').hidden = (DUAL_TRACK.path !== 'PATH_B');
  if (DUAL_TRACK.path === 'PATH_B') bindCaseInput();

  /* 步驟 04 / 05 */
  var model = pyOk ? getModel() : null;
  $('stepFilter').hidden = !DUAL_TRACK.path;
  $('filterHint').textContent = (DUAL_TRACK.path === 'PATH_A')
    ? '決定查詢表的查詢範圍，留空＝整份主檔'
    : '僅影響四分頁 Excel，不影響提案 Word';

  $('stepOutA').hidden = (DUAL_TRACK.path !== 'PATH_A');
  $('stepOutB').hidden = (DUAL_TRACK.path !== 'PATH_B');

  if (DUAL_TRACK.path === 'PATH_A'){
    $('btnQuery').disabled = !model;
  }
  if (DUAL_TRACK.path === 'PATH_B'){
    var ready = !!(model && DUAL_TRACK.pathB.caseData && pathBLoaded());
    ['btnPreview','btnDocx','btnAttach'].forEach(function(id){ if ($(id)) $(id).disabled = !ready; });
    if ($('btnXlsx')) $('btnXlsx').disabled = !(model && pathBLoaded());
    if ($('btnTpl')) $('btnTpl').disabled = !pathBLoaded();
  }
}

/* ============================================================
   路徑 A：產出查詢表
   ============================================================ */
function currentFilters(){
  return {atc7:parseFilter($('qAtc7').value), grpCd:parseFilter($('qGrpCd').value),
          grpNm:parseFilter($('qGrpNm').value), code:parseFilter($('qCode').value),
          brand:parseFilter($('qBrand').value)};
}
$('btnQuery').addEventListener('click', function(){
  var model = getModel();
  if (!model){ log('請先載入主檔並選定 price 年月。', 'bad'); return; }
  var py = currentPy();
  try {
    var wb = buildQueryWorkbook(model, currentFilters(), py);
    DUAL_TRACK.pathA.sheets = {sheet1:wb.counts.s1, sheet2:wb.counts.s2};
    if (!wb.counts.sel){ log('查無符合篩選條件的項目，查詢表仍會產出（僅含表頭）。', 'warn'); }
    log('查詢表：符合條件 ' + wb.counts.sel.toLocaleString() + ' 筆 → 定位 ' + wb.counts.groups
        + ' 個分組，展開 ' + wb.counts.s1 + ' 項；分組×ATC彙總 ' + wb.counts.s2 + ' 列', 'ok');
    save(wb.book, '藥品背景查詢_' + py + '.xlsx');
  } catch (err){ log('查詢表產出失敗：' + err.message, 'bad'); }
});

/* ============================================================
   路徑 B：預覽與產出（模組載入後才綁定）
   ============================================================ */
function buildReport(){
  var py = currentPy();
  var model = getModel();
  var cs = DUAL_TRACK.pathB.caseData;
  if (!model || !cs) return null;
  try {
    var rep = buildCaseReport(cs, model, py);
    if (rep.missing.length) log('主檔查無這些 CODE，已略過：' + rep.missing.join('、'), 'warn');
    DUAL_TRACK.pathB.targets = rep;
    return rep;
  } catch (err){ log('產生失敗：' + err.message, 'bad'); return null; }
}
function selectedAtcLevels(){
  var out = [];
  [1,3,4,5,7].forEach(function(n){ if ($('atcL'+n) && $('atcL'+n).checked) out.push(n); });
  return out;
}
function renderTable(head, rows){
  var h = '<div class="scroll"><table><thead><tr>';
  for (var i=0;i<head.length;i++) h += '<th>'+esc(head[i])+'</th>';
  h += '</tr></thead><tbody>';
  for (var r=0;r<rows.length;r++){
    h += '<tr'+(rows[r].__self?' class="self"':'')+'>';
    for (i=0;i<head.length;i++){
      var v = rows[r][i] === undefined ? '' : rows[r][i];
      if (isSeg(v)) v = segText(v);
      var isNum = typeof v === 'string' && /^[\d,.\s%－-]*$/.test(v) && v !== '';
      h += '<td'+(isNum?' class="num"':'')+'>'+esc(v)+'</td>';
    }
    h += '</tr>';
  }
  return h + '</tbody></table></div>';
}

var _pathBBound = false;
function bindPathBHandlers(){
  if (_pathBBound) return;
  _pathBBound = true;

  $('btnTpl').addEventListener('click', function(){
    try {
      save(buildXlsx(TPL, '案件需求', {validations: TPL_VALID, extraSheets: [tplHelpSheet()]}),
           '案件需求檔_範本.xlsx',
           {blob: csvBlob(TPL), name: '案件需求檔_範本.csv'});
      log('範本已下載：案件需求檔_範本.xlsx（' + TPL_KEYS.length + ' 個欄位）', 'ok');
    } catch (err){ log('範本產生失敗：' + err.message, 'bad'); }
  });

  $('btnPreview').addEventListener('click', function(){
    var rep = buildReport(), box = $('prevBox');
    box.innerHTML = '';
    if (!rep) return;
    var el = document.createElement('div');
    el.className = 'case-card';
    var h = '<div class="case-hd"><b>本案 ' + rep.items.length + ' 項藥品</b>'
          + '<code>報告人 ' + esc(rep.reporter) + '</code>'
          + '<code>申報 ' + rep.dataYear + ' 年</code></div><div class="case-bd">';
    var CN = '一二三四五六七八九十', ci = rep.items.length;
    h += '<div><div class="tbl-cap">（' + CN[ci++] + '）查健保收載同分組藥品共計 ' + rep.table1.length + ' 項(P-)（已排除現行支付價 0 元者）</div>'
       + renderTable(rep.table1Head, rep.table1) + '</div>';
    for (var bi=0;bi<rep.table2Blocks.length;bi++){
      var bk = rep.table2Blocks[bi];
      h += '<div><div class="tbl-cap">（' + CN[ci++] + '）查健保收載' + (bk.mode === 'PREFIX5' ? '相同ATC 5碼' : '相同ATC 7碼')
         + '藥品共計 ' + bk.nGroups + ' 分組、' + bk.nItems + ' 項目(P-)</div>'
         + renderTable(bk.head, bk.rows)
         + bk.notes.map(function(n){ return '<div class="tbl-note">' + esc(n) + '</div>'; }).join('')
         + '</div>';
    }
    h += '<div><div class="tbl-cap">五、廠商建議(P-)</div>' + renderTable(rep.table5Head, rep.table5) + '</div>';
    for (var i=0;i<rep.items.length;i++){
      var it = rep.items[i], c = it.calc;
      var head = ['項目'].concat(c.tenHead.map(function(x){ return '十國 '+x; }))
                          .concat(c.costHead.map(function(x){ return '成本 '+x; }));
      var rows = [
        ['調高後支付價'].concat(c.priceRow),
        ['整體藥費'].concat(c.totalRow),
        ['財務衝擊'].concat(c.impactRow)
      ];
      h += '<div><div class="tbl-cap">六、（' + '一二三四五六七八九十'[i] + '）'
         + esc(txt(it.drug['藥品名稱'])) + '（' + esc(it.code) + '）　健保價 '
         + (it.row.price === null ? '－' : money(it.row.price))
         + '　建議價 ' + (it.sug === null ? '－' : money(it.sug))
         + '　近三年平均量 ' + (c.avgQty === null ? '－' : fmt(Math.round(c.avgQty)))
         + '</div>' + renderTable(head, rows) + '</div>';
    }
    var nt = '';
    for (i=0;i<rep.notes.length;i++){
      var n = rep.notes[i];
      var tag = rep.notes.length > 1 ? '（項次 ' + n.idx.join('、') + '）' : '';
      nt += (nt ? '\n' : '') + '註1：' + tag + n.a10Law + '\n註2：' + tag + n.costLaw
          + '\n　　2.領有藥物許可證者，得加計繳納藥害救濟徵收金比率及營業稅。';
      for (var q=0;q<n.licOrder.length;q++){
        var lt = n.licOrder.length > 1 ? '（項次 ' + n.lics[n.licOrder[q]].join('、') + '）' : '';
        nt += '\n　　　' + lt + n.licOrder[q];
      }
    }
    h += '<div class="notes">' + esc(nt) + '\n' + esc(rep.art40) + '</div>';
    box.appendChild(el);
    el.innerHTML = h + '</div>';
    log('預覽完成：' + rep.items.length + ' 項藥品、表一 ' + rep.table1.length + ' 列。', 'ok');
  });

  $('btnXlsx').addEventListener('click', function(){
    var model = getModel();
    if (!model){ log('請先載入藥品主檔並選定 price 年月。', 'bad'); return; }
    var py = currentPy();
    if (!DUAL_TRACK.pathB.caseData){
      log('已選正式提案，請先上傳案件需求檔（分頁 1 將只輸出說明）。', 'warn');
    }
    var rep0 = buildReport();
    var targets = rep0 ? rep0.items : [];
    try {
      var wb = buildSasWorkbook(model, currentFilters(), targets, py, {skipPriceCalc:false});
      DUAL_TRACK.pathB.outputs.xlsxMain = wb.counts;
      log('篩選後 ' + wb.counts.sel.toLocaleString() + ' 筆 → 分組統計 ' + wb.counts.s2
          + ' 列、明細 ' + wb.counts.s3 + ' 列、含0元 ' + wb.counts.s4 + ' 列', 'ok');
      if (wb.counts.s4 > 60000) log('資料量偏大，Excel 產出可能需要數十秒。', 'warn');
      save(wb.book, '藥價核價報表_' + py + '.xlsx');
    } catch (err){ log('Excel 產出失敗：' + err.message, 'bad'); }
  });

  $('btnAttach').addEventListener('click', function(){
    var model = getModel(), rep = buildReport();
    if (!model || !rep) return;
    var py = currentPy();
    try {
      var at = buildAttachmentWorkbook(model, rep, py, selectedAtcLevels());
      DUAL_TRACK.pathB.outputs.xlsxAttach = at.counts;
      save(at.book, '提案附件_' + py + '.xlsx');
      log('已產出提案附件 Excel：' + at.counts.join('、'), 'ok');
    } catch (err){ log('提案附件產出失敗：' + err.message, 'bad'); }
  });

  $('btnDocx').addEventListener('click', function(){
    var rep = buildReport();
    if (!rep) return;
    try {
      var first = txt(rep.items[0].drug['藥品名稱']).replace(/[\\/:*?"<>|]/g,'').substring(0,20);
      var name = '提案_' + first + (rep.items.length > 1 ? '等' + rep.items.length + '項' : '') + '.docx';
      DUAL_TRACK.pathB.outputs.word = name;
      save(reportToDocx(rep), name);
      log('已產出 ' + name + '（' + rep.items.length + ' 項藥品）', 'ok');
    } catch (err){ log('產出失敗：' + err.message, 'bad'); }
  });
}

refresh();
