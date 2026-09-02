'use strict';
/* DOCX OOXML 引擎（僅 Path B 載入） */

var FONTS = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="標楷體" w:cs="Times New Roman"/>';

var SZ    = 28;          /* 14pt */
var LINE  = 480;         /* 固定行高 24pt */

function runs(text, opt){
  opt = opt || {};
  var sz = opt.sz || SZ;
  var rpr = '<w:rPr>' + FONTS + (opt.b ? '<w:b/>' : '')
    + '<w:sz w:val="'+sz+'"/><w:szCs w:val="'+sz+'"/></w:rPr>';
  var lines = String(text === undefined || text === null ? '' : text).split('\n');
  var out = '';
  for (var i=0;i<lines.length;i++){
    if (i) out += '<w:r>'+rpr+'<w:br/></w:r>';
    out += '<w:r>'+rpr+'<w:t xml:space="preserve">'+esc(lines[i])+'</w:t></w:r>';
  }
  return out;
}

/* w:pPr 子元素順序須遵守 schema：numPr → tabs → spacing → ind → jc → rPr */
function pPr(opt){
  opt = opt || {};
  var sz = opt.sz || SZ;
  var x = '<w:pPr>';
  if (opt.numId) x += '<w:numPr><w:ilvl w:val="'+(opt.ilvl||0)+'"/><w:numId w:val="'+opt.numId+'"/></w:numPr>';
  if (opt.tabRight) x += '<w:tabs><w:tab w:val="right" w:pos="'+opt.tabRight+'"/></w:tabs>';
  x += '<w:spacing w:before="'+(opt.before||0)+'" w:after="'+(opt.after===undefined?0:opt.after)+'"'
     + ' w:line="'+(opt.line || LINE)+'" w:lineRule="'+(opt.lineRule||'exact')+'"/>';
  if (opt.indent !== undefined) x += '<w:ind w:left="'+opt.indent+'"'+(opt.hanging?' w:hanging="'+opt.hanging+'"':'')+'/>';
  if (opt.align) x += '<w:jc w:val="'+opt.align+'"/>';
  x += '<w:rPr>'+FONTS+(opt.b?'<w:b/>':'')+'<w:sz w:val="'+sz+'"/><w:szCs w:val="'+sz+'"/></w:rPr>';
  return x + '</w:pPr>';
}

/* 同一段落中混排不同格式的文字段：[{t:'…', b:true, sup:true, sz:22}, …] */
function segRuns(segs, opt){
  opt = opt || {};
  var base = opt.sz || SZ, out = '';
  for (var i=0;i<segs.length;i++){
    var s = segs[i] || {}, sz = s.sz || base;
    var rpr = '<w:rPr>' + FONTS + (s.b ? '<w:b/>' : '')
      + '<w:sz w:val="'+sz+'"/><w:szCs w:val="'+sz+'"/>'
      + (s.sup ? '<w:vertAlign w:val="superscript"/>' : '') + '</w:rPr>';
    var lines = String(s.t === undefined || s.t === null ? '' : s.t).split('\n');
    for (var j=0;j<lines.length;j++){
      if (j) out += '<w:r>'+rpr+'<w:br/></w:r>';
      out += '<w:r>'+rpr+'<w:t xml:space="preserve">'+esc(lines[j])+'</w:t></w:r>';
    }
  }
  return out;
}

function paraSeg(segs, opt){ return '<w:p>' + pPr(opt) + segRuns(segs, opt) + '</w:p>'; }

/* 儲存格值可為字串，或 {seg:[…]} 以混排格式 */
function segCell(list){ return {seg:list}; }

function isSeg(v){ return !!(v && typeof v === 'object' && v.seg); }

function segText(v){
  if (!isSeg(v)) return v === undefined || v === null ? '' : String(v);
  return v.seg.map(function(s){ return s.t === undefined ? '' : String(s.t); }).join('');
}

function para(text, opt){ return '<w:p>' + pPr(opt) + runs(text, opt) + '</w:p>'; }

function paraTab(left, right, opt){
  opt = opt || {}; opt.tabRight = 9020;
  var sz = opt.sz || SZ;
  var rpr = '<w:rPr>'+FONTS+(opt.b?'<w:b/>':'')+'<w:sz w:val="'+sz+'"/><w:szCs w:val="'+sz+'"/></w:rPr>';
  return '<w:p>' + pPr(opt)
    + '<w:r>'+rpr+'<w:t xml:space="preserve">'+esc(left)+'</w:t></w:r>'
    + '<w:r>'+rpr+'<w:tab/><w:t xml:space="preserve">'+esc(right)+'</w:t></w:r>'
    + '</w:p>';
}

/* ---- 表格 ---- */
var BORDERS = '<w:tblBorders>'
  + ['top','left','bottom','right','insideH','insideV'].map(function(s){
      return '<w:'+s+' w:val="single" w:sz="6" w:space="0" w:color="000000"/>';
    }).join('')
  + '</w:tblBorders>';

var SHADE_HD = 'D9D9D9';

function tblOpen(w, totalW){
  var xml = '<w:tbl><w:tblPr><w:tblW w:w="'+totalW+'" w:type="dxa"/>'
          + '<w:jc w:val="center"/>' + BORDERS
          + '<w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/>'
          + '<w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar>'
          + '</w:tblPr><w:tblGrid>';
  for (var i=0;i<w.length;i++) xml += '<w:gridCol w:w="'+w[i]+'"/>';
  return xml + '</w:tblGrid>';
}

function widthsOf(weights, totalW){
  var sum = weights.reduce(function(a,b){return a+b;},0);
  return weights.map(function(x){ return Math.round(x/sum*totalW); });
}

function tcell(txt, wid, opt){
  opt = opt || {};
  var pr = '<w:tcPr><w:tcW w:w="'+wid+'" w:type="dxa"/>';
  if (opt.span > 1) pr += '<w:gridSpan w:val="'+opt.span+'"/>';
  if (opt.vm === 'restart') pr += '<w:vMerge w:val="restart"/>';
  else if (opt.vm === 'cont') pr += '<w:vMerge/>';
  if (opt.shade) pr += '<w:shd w:val="clear" w:color="auto" w:fill="'+opt.shade+'"/>';
  pr += '<w:vAlign w:val="center"/></w:tcPr>';
  var pOpt = {sz:opt.sz||24, b:opt.b, align:opt.align, after:0, line:opt.line||320, lineRule:'auto'};
  var body;
  if (opt.vm === 'cont') body = para('', pOpt);
  else if (isSeg(txt))   body = paraSeg(txt.seg, pOpt);
  else                   body = para(txt, pOpt);
  return '<w:tc>' + pr + body + '</w:tc>';
}

var NUMLIKE = /^[\d,.\s%－\-億萬元]*$/;

/* 一般表格：head 為字串陣列，rows 為值陣列 */
function docTable(head, rows, weights, totalW, sz){
  totalW = totalW || 9020;
  var w = widthsOf(weights, totalW), i;
  var xml = tblOpen(w, totalW);
  xml += '<w:tr><w:trPr><w:tblHeader/></w:trPr>';
  for (i=0;i<head.length;i++) xml += tcell(head[i], w[i], {b:true, align:'center', shade:SHADE_HD, sz:sz});
  xml += '</w:tr>';
  for (var r=0;r<rows.length;r++){
    xml += '<w:tr>';
    for (i=0;i<head.length;i++){
      var v = rows[r][i];
      var isNum = typeof v === 'string' && NUMLIKE.test(v) && v !== '';
      /* 資料列一律不加粗；只有 {seg:[…]} 內個別文字段可自帶粗體 */
      xml += tcell(v, w[i], {sz:sz,
        align: rows[r].__center ? 'center' : (isNum ? 'right' : (isSeg(v) ? 'center' : 'left'))});
    }
    xml += '</w:tr>';
  }
  return xml + '</w:tbl>';
}

/* 進階表格：rows = [[{t,span,vm,b,align,shade}, …], …]，可合併儲存格 */
function docTableRaw(rows, weights, totalW, sz){
  totalW = totalW || 9020;
  var w = widthsOf(weights, totalW);
  var xml = tblOpen(w, totalW);
  for (var r=0;r<rows.length;r++){
    xml += '<w:tr>' + (r === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : '');
    var col = 0;
    for (var i=0;i<rows[r].length;i++){
      var c = rows[r][i] || {};
      var span = c.span || 1, wid = 0;
      for (var k=0;k<span;k++) wid += (w[col+k] || 0);
      xml += tcell(c.t === undefined ? '' : c.t, wid, {
        span: span, vm: c.vm, b: c.b, shade: c.shade,
        align: c.align || 'center', sz: c.sz || sz
      });
      col += span;
    }
    xml += '</w:tr>';
  }
  return xml + '</w:tbl>';
}

/* ---- 章節與封裝 ---- */
var SECT_PORTRAIT  = '<w:pgSz w:w="11906" w:h="16838"/>'
  + '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>'
  + '<w:docGrid w:type="lines" w:linePitch="360"/>';

var SECT_LANDSCAPE = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
  + '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>'
  + '<w:docGrid w:type="lines" w:linePitch="360"/>';

var LAND_W = 16838 - 1418 - 1134;

function sectBreak(sect){ return '<w:p><w:pPr><w:sectPr>' + sect + '</w:sectPr></w:pPr></w:p>'; }

/* 多階編號：一、 →（一）→ 1. */
var NUM_ID = 1;

function numberingXml(){
  function lvl(i, fmt, text, left, hang){
    return '<w:lvl w:ilvl="'+i+'"><w:start w:val="1"/><w:numFmt w:val="'+fmt+'"/>'
      + '<w:lvlText w:val="'+esc(text)+'"/><w:lvlJc w:val="left"/>'
      + '<w:pPr><w:ind w:left="'+left+'" w:hanging="'+hang+'"/></w:pPr>'
      + '<w:rPr>'+FONTS+'</w:rPr></w:lvl>';
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>'
    /* 縮排以 14pt 中文字寬 280 twips 為單位：
       第一層「一、」起排於 1 個字元（280）、第二層「（一）」起排於 1.5 個字元（420） */
    + lvl(0, 'chineseCounting', '%1、',   840,  560)   /* 280 + 標籤 2 字 */
    + lvl(1, 'chineseCounting', '（%2）', 1260, 840)   /* 420 + 標籤 3 字 */
    + lvl(2, 'decimal',         '%3.',   1680, 420)
    + lvl(3, 'decimal',         '(%4)',  2100, 420)
    + '</w:abstractNum>'
    + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    + '</w:numbering>';
}

function buildDocx(bodyXml){
  var doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:body>' + bodyXml
    + '<w:sectPr>' + SECT_PORTRAIT + '</w:sectPr></w:body></w:document>';
  var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:docDefaults><w:rPrDefault><w:rPr>'+FONTS+'<w:sz w:val="'+SZ+'"/><w:szCs w:val="'+SZ+'"/></w:rPr></w:rPrDefault>'
    + '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="'+LINE+'" w:lineRule="exact"/></w:pPr></w:pPrDefault>'
    + '</w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
    + '<w:qFormat/></w:style>'
    + '</w:styles>';
  var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    + '</Types>';
  var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '</Relationships>';
  var drels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    + '</Relationships>';
  return zipWrite([
    {name:'[Content_Types].xml',          data:TE.encode(ct)},
    {name:'_rels/.rels',                  data:TE.encode(rels)},
    {name:'word/document.xml',            data:TE.encode(doc)},
    {name:'word/styles.xml',              data:TE.encode(styles)},
    {name:'word/numbering.xml',           data:TE.encode(numberingXml())},
    {name:'word/_rels/document.xml.rels', data:TE.encode(drels)}
  ]);
}
