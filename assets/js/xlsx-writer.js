'use strict';
/* XLSX 組出引擎（Path A / Path B 共用） */

function xlColName(i){ var s=''; i++; while(i>0){ var m=(i-1)%26; s=String.fromCharCode(65+m)+s; i=(i-m-1)/26; } return s; }

/* 樣式索引（對應下方 cellXfs 順序） */
var XS = {
  BODY:0, HEAD:1, TITLE:2, INT:3, DEC:4, SUB:5, NOTE:6, MARK:7, LABEL:8, SUBINT:9, SECT:10
};

function xcell(v, s){ return {v:v, s:s}; }

function stylesXml(){
  var F = '"Microsoft JhengHei"';
  function font(sz, b, color, name){
    return '<font><sz val="'+sz+'"/>' + (b ? '<b/>' : '')
      + '<color rgb="FF'+color+'"/><name val="'+(name||'Microsoft JhengHei')+'"/><charset val="136"/></font>';
  }
  function fill(rgb){
    return rgb ? '<fill><patternFill patternType="solid"><fgColor rgb="FF'+rgb+'"/><bgColor indexed="64"/></patternFill></fill>'
               : '<fill><patternFill patternType="none"/></fill>';
  }
  var thin = '<left style="thin"><color rgb="FFC8C8C8"/></left>'
           + '<right style="thin"><color rgb="FFC8C8C8"/></right>'
           + '<top style="thin"><color rgb="FFC8C8C8"/></top>'
           + '<bottom style="thin"><color rgb="FFC8C8C8"/></bottom><diagonal/>';
  function xf(fontId, fillId, borderId, numFmt, align){
    return '<xf numFmtId="'+(numFmt||0)+'" fontId="'+fontId+'" fillId="'+fillId+'" borderId="'+borderId+'" xfId="0"'
      + (numFmt ? ' applyNumberFormat="1"' : '') + ' applyFont="1" applyFill="1" applyBorder="1"'
      + (align ? ' applyAlignment="1"><alignment '+align+'/></xf>' : '/>');
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.###"/></numFmts>'
    + '<fonts count="5">'
      + font(10, false, '161616')          /* 0 內文 */
      + font(10, true,  '161616')          /* 1 粗體 */
      + font(14, true,  '0F62FE')          /* 2 標題 */
      + font(9,  false, '262626')          /* 3 備註 */
      + font(11, true,  '161616')          /* 4 區塊標題 */
    + '</fonts>'
    + '<fills count="5">' + fill(null) + '<fill><patternFill patternType="gray125"/></fill>'
      + fill('D0E2FF')                     /* 2 表頭淺藍 */
      + fill('F4F4F4')                     /* 3 小計淺灰 */
      + fill('EDF5FF')                     /* 4 標記淡藍 */
    + '</fills>'
    + '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
      + '<border>' + thin + '</border></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="11">'
      + xf(0,0,1,0,'vertical="center" wrapText="1"')                        /* 0 BODY */
      + xf(1,2,1,0,'horizontal="center" vertical="center" wrapText="1"')    /* 1 HEAD */
      + xf(2,0,0,0,'vertical="center"')                                     /* 2 TITLE */
      + xf(0,0,1,3,'horizontal="right" vertical="center"')                  /* 3 INT  #,##0 */
      + xf(0,0,1,164,'horizontal="right" vertical="center"')                /* 4 DEC  #,##0.### */
      + xf(1,3,1,0,'vertical="center" wrapText="1"')                        /* 5 SUB  小計 */
      + xf(3,0,0,0,'vertical="center"')                                     /* 6 NOTE 不換行，可溢出 */
      + xf(1,4,1,0,'vertical="center" wrapText="1"')                        /* 7 MARK 本案 */
      + xf(1,3,1,0,'horizontal="center" vertical="center" wrapText="1"')    /* 8 LABEL 左欄 */
      + xf(1,3,1,3,'horizontal="right" vertical="center"')                  /* 9 SUBINT 小計數字 */
      + xf(4,0,0,0,'vertical="center"')                                     /* 10 SECT 區塊標題 */
    + '</cellXfs>'
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    + '</styleSheet>';
}

function sheetXml(sh){
  var aoa = sh.aoa || [], i, r;
  var maxCol = 0;
  for (r=0;r<aoa.length;r++) if (aoa[r] && aoa[r].length > maxCol) maxCol = aoa[r].length;

  var out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'];
  if (sh.fitToPage) out.push('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
  out.push('<dimension ref="A1:' + xlColName(Math.max(maxCol,1)-1) + Math.max(aoa.length,1) + '"/>');

  /* 凍結窗格 */
  var fr = sh.freeze || 0;
  if (fr > 0){
    out.push('<sheetViews><sheetView workbookViewId="0">'
      + '<pane ySplit="'+fr+'" topLeftCell="A'+(fr+1)+'" activePane="bottomLeft" state="frozen"/>'
      + '<selection pane="bottomLeft" activeCell="A'+(fr+1)+'" sqref="A'+(fr+1)+'"/>'
      + '</sheetView></sheetViews>');
  } else {
    out.push('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
  }
  out.push('<sheetFormatPr defaultRowHeight="' + (sh.rowHeight || 20) + '" defaultColWidth="12"/>');

  /* 欄寬 */
  if (sh.cols && sh.cols.length){
    var c = '<cols>';
    for (i=0;i<sh.cols.length;i++)
      c += '<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+sh.cols[i]+'" customWidth="1"/>';
    out.push(c + '</cols>');
  }

  out.push('<sheetData>');
  for (r=0;r<aoa.length;r++){
    var vals = aoa[r] || [], n = r+1;
    var ht = sh.rowHeights && sh.rowHeights[r];
    var row = '<row r="'+n+'"' + (ht ? ' ht="'+ht+'" customHeight="1"' : '') + '>';
    for (i=0;i<vals.length;i++){
      var cell = vals[i];
      if (cell === undefined || cell === null) continue;
      var isObj = (typeof cell === 'object' && cell !== null && 'v' in cell);
      var v  = isObj ? cell.v : cell;
      var st = (typeof cell === 'object' && cell !== null && 's' in cell) ? cell.s : null;
      var empty = (v === undefined || v === null || v === '');
      /* 空白但有指定樣式者仍要寫出（只帶 s、不帶值），
         否則小計列的灰底與表格框線會在空欄位處斷掉 */
      if (empty && st === null) continue;
      if (st === null) st = (typeof v === 'number' && isFinite(v)) ? XS.INT : XS.BODY;
      var ref = xlColName(i)+n, sa = ' s="'+st+'"';
      if (empty) row += '<c r="'+ref+'"'+sa+'/>';
      else if (typeof v === 'number' && isFinite(v)) row += '<c r="'+ref+'"'+sa+'><v>'+v+'</v></c>';
      else row += '<c r="'+ref+'"'+sa+' t="inlineStr"><is><t xml:space="preserve">'+esc(v)+'</t></is></c>';
    }
    out.push(row+'</row>');
  }
  out.push('</sheetData>');

  if (sh.autoFilter) out.push('<autoFilter ref="'+sh.autoFilter+'"/>');
  if (sh.merges && sh.merges.length){
    out.push('<mergeCells count="'+sh.merges.length+'">'
      + sh.merges.map(function(m){ return '<mergeCell ref="'+m+'"/>'; }).join('')
      + '</mergeCells>');
  }
  if (sh.validations && sh.validations.length){
    out.push('<dataValidations count="'+sh.validations.length+'">'
      + sh.validations.map(function(dv){
          return '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"'
               + ' promptTitle="'+esc(dv.title||'')+'" prompt="'+esc(dv.prompt||'')+'"'
               + ' errorTitle="輸入值不符" error="請由清單選取" sqref="'+dv.ref+'">'
               + '<formula1>&quot;'+esc(dv.list.join(','))+'&quot;</formula1></dataValidation>';
        }).join('')
      + '</dataValidations>');
  }
  out.push('<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>');
  if (sh.landscape || sh.fitToPage){
    out.push('<pageSetup paperSize="9"'                      /* 9 = A4 */
      + (sh.landscape ? ' orientation="landscape"' : '')
      + (sh.fitToPage ? ' fitToWidth="1" fitToHeight="0"' : '') + '/>');
  }
  out.push('</worksheet>');
  return out.join('');
}

/* 依內容自動估欄寬（全形字算 2） */
function autoCols(aoa, min, max){
  min = min || 8; max = max || 42;
  var w = [], r, i;
  for (r=0;r<aoa.length;r++){
    var row = aoa[r] || [];
    for (i=0;i<row.length;i++){
      var cell = row[i];
      var v = (cell && typeof cell === 'object' && 'v' in cell) ? cell.v : cell;
      if (v === undefined || v === null) continue;
      var t = String(typeof v === 'number' ? Math.round(v).toLocaleString('en-US') : v);
      var longest = 0;
      t.split('\n').forEach(function(line){
        var n = 0;
        for (var k=0;k<line.length;k++) n += (line.charCodeAt(k) > 255 ? 2 : 1);
        if (n > longest) longest = n;
      });
      if (!w[i] || longest > w[i]) w[i] = longest;
    }
  }
  for (i=0;i<w.length;i++) w[i] = Math.min(max, Math.max(min, (w[i]||min) + 2));
  return w;
}

function rowsToAoa(rows){
  var head = Object.keys(rows[0] || {});
  return [head.map(function(h){ return xcell(h, XS.HEAD); })]
    .concat(rows.map(function(r){ return head.map(function(h){ return r[h]; }); }));
}

function buildXlsxBook(sheets){
  var files = [], sheetRels = [], sheetTags = [], defNames = [], i;
  for (i=0;i<sheets.length;i++){
    var sh = sheets[i];
    if (!sh.aoa) sh.aoa = rowsToAoa(sh.rows || []);
    if (!sh.cols) sh.cols = autoCols(sh.aoa);
    var nm = String(sh.name).replace(/[\\\/\?\*\[\]:]/g,'_').substring(0,31);
    files.push({name:'xl/worksheets/sheet'+(i+1)+'.xml', data:TE.encode(sheetXml(sh))});
    sheetRels.push('<Relationship Id="rId'+(i+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+(i+1)+'.xml"/>');
    sheetTags.push('<sheet name="'+esc(nm)+'" sheetId="'+(i+1)+'" r:id="rId'+(i+1)+'"/>');
    var qn = "'" + nm.replace(/'/g,"''") + "'";
    if (sh.printArea)
      defNames.push('<definedName name="_xlnm.Print_Area" localSheetId="'+i+'">'+esc(qn+'!'+sh.printArea)+'</definedName>');
    if (sh.printTitles)
      defNames.push('<definedName name="_xlnm.Print_Titles" localSheetId="'+i+'">'+esc(qn+'!'+sh.printTitles)+'</definedName>');
  }
  var styleRid = 'rId' + (sheets.length + 1);
  var wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<sheets>' + sheetTags.join('') + '</sheets>'
    + (defNames.length ? '<definedNames>' + defNames.join('') + '</definedNames>' : '')
    + '</workbook>';
  var wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + sheetRels.join('')
    + '<Relationship Id="'+styleRid+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    + '</Relationships>';
  var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + sheets.map(function(_,k){ return '<Override PartName="/xl/worksheets/sheet'+(k+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'; }).join('')
    + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    + '</Types>';
  var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    + '</Relationships>';
  return zipWrite([
    {name:'[Content_Types].xml',       data:TE.encode(ct)},
    {name:'_rels/.rels',               data:TE.encode(rels)},
    {name:'xl/workbook.xml',           data:TE.encode(wb)},
    {name:'xl/_rels/workbook.xml.rels',data:TE.encode(wbRels)},
    {name:'xl/styles.xml',             data:TE.encode(stylesXml())}
  ].concat(files));
}

function buildXlsx(rows, sheetName, opt){
  var sh = {name:sheetName, rows:rows, freeze:1}, k;
  if (opt) for (k in opt) sh[k] = opt[k];
  return buildXlsxBook([sh].concat((opt && opt.extraSheets) || []));
}

/* 依欄名決定數值格式 */
function styleForHeader(h){
  h = String(h || '');
  if (/金額|數量|項目數|品項數|總數|申報量/.test(h)) return XS.INT;
  if (/支付價|價格|中位價|成本|建議價/.test(h)) return XS.DEC;
  return XS.BODY;
}

function headRow(arr){ return arr.map(function(h){ return xcell(h, XS.HEAD); }); }

function dataRow(vals, styles, override){
  return vals.map(function(v, i){
    var st = (override && override[i] !== undefined) ? override[i] : styles[i];
    if (typeof v !== 'number' && st !== XS.BODY && st !== XS.SUB && st !== XS.MARK) st = XS.BODY;
    return xcell(v, st);
  });
}
