'use strict';
/* ZIP 讀寫 + XLSX / CSV 解析（Path A / Path B 共用） */

var CRC = (function(){
  var t = new Uint32Array(256);
  for (var n=0;n<256;n++){ var c=n;
    for (var k=0;k<8;k++) c = (c&1) ? (0xEDB88320 ^ (c>>>1)) : (c>>>1);
    t[n]=c>>>0; }
  return t;
})();

function crc32(b){
  var c = 0xFFFFFFFF;
  for (var i=0;i<b.length;i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c>>>8);
  return (c ^ 0xFFFFFFFF)>>>0;
}

var TE = new TextEncoder();

var TD = new TextDecoder('utf-8');

/* ---- 寫：有 CompressionStream 就 deflate，否則 stored ---- */
function deflateRaw(bytes){
  if (typeof CompressionStream === 'undefined') return Promise.resolve(null);
  var s = new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw'));
  return new Response(s).arrayBuffer().then(function(a){ return new Uint8Array(a); })
                        .catch(function(){ return null; });
}

function zipWrite(files){
  return Promise.all(files.map(function(f){
    if (f.data.length < 256) return Promise.resolve({f:f, z:null});
    return deflateRaw(f.data).then(function(z){
      return {f:f, z:(z && z.length < f.data.length) ? z : null};
    });
  })).then(function(list){
    var parts=[], central=[], offset=0, i;
    for (i=0;i<list.length;i++){
      var f = list[i].f, z = list[i].z;
      var method = z ? 8 : 0, payload = z || f.data;
      var nameB = TE.encode(f.name), crc = crc32(f.data);
      var lh = new Uint8Array(30+nameB.length), lv = new DataView(lh.buffer);
      lv.setUint32(0,0x04034b50,true); lv.setUint16(4,20,true); lv.setUint16(6,0x0800,true);
      lv.setUint16(8,method,true); lv.setUint16(10,0,true); lv.setUint16(12,0,true);
      lv.setUint32(14,crc,true); lv.setUint32(18,payload.length,true); lv.setUint32(22,f.data.length,true);
      lv.setUint16(26,nameB.length,true); lv.setUint16(28,0,true);
      lh.set(nameB,30);
      parts.push(lh,payload);

      var ch = new Uint8Array(46+nameB.length), cv = new DataView(ch.buffer);
      cv.setUint32(0,0x02014b50,true); cv.setUint16(4,20,true); cv.setUint16(6,20,true);
      cv.setUint16(8,0x0800,true); cv.setUint16(10,method,true);
      cv.setUint16(12,0,true); cv.setUint16(14,0,true);
      cv.setUint32(16,crc,true); cv.setUint32(20,payload.length,true); cv.setUint32(24,f.data.length,true);
      cv.setUint16(28,nameB.length,true); cv.setUint16(30,0,true); cv.setUint16(32,0,true);
      cv.setUint16(34,0,true); cv.setUint16(36,0,true); cv.setUint32(38,0,true);
      cv.setUint32(42,offset,true);
      ch.set(nameB,46);
      central.push(ch);
      offset += lh.length + payload.length;
    }
    var cdSize=0;
    for (i=0;i<central.length;i++) cdSize += central[i].length;
    var eo = new Uint8Array(22), ev = new DataView(eo.buffer);
    ev.setUint32(0,0x06054b50,true); ev.setUint16(4,0,true); ev.setUint16(6,0,true);
    ev.setUint16(8,central.length,true); ev.setUint16(10,central.length,true);
    ev.setUint32(12,cdSize,true); ev.setUint32(16,offset,true); ev.setUint16(20,0,true);
    return new Blob(parts.concat(central,[eo]),{type:'application/octet-stream'});
  });
}

/* ---- 讀：走中央目錄，deflate 用瀏覽器內建解壓 ---- */
function inflateRaw(bytes){
  if (typeof DecompressionStream === 'undefined')
    return Promise.reject(new Error('此瀏覽器不支援 DecompressionStream，請改用 Chrome / Edge 80+，或改上傳 CSV'));
  var s = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Response(s).arrayBuffer().then(function(a){ return new Uint8Array(a); });
}

function zipRead(buf){
  var u = new Uint8Array(buf), dv = new DataView(buf), i;
  var eocd = -1;
  for (i = u.length-22; i >= 0 && i >= u.length-22-65557; i--){
    if (dv.getUint32(i,true) === 0x06054b50){ eocd = i; break; }
  }
  if (eocd < 0) throw new Error('不是有效的 xlsx（找不到 ZIP 結尾）');
  var count = dv.getUint16(eocd+10,true), cdOff = dv.getUint32(eocd+16,true);
  var entries = [], p = cdOff;
  for (i=0;i<count;i++){
    if (dv.getUint32(p,true) !== 0x02014b50) break;
    var method = dv.getUint16(p+10,true);
    var cSize  = dv.getUint32(p+20,true);
    var nLen   = dv.getUint16(p+28,true);
    var eLen   = dv.getUint16(p+30,true);
    var kLen   = dv.getUint16(p+32,true);
    var lOff   = dv.getUint32(p+42,true);
    var name   = TD.decode(u.subarray(p+46, p+46+nLen));
    entries.push({name:name,method:method,cSize:cSize,lOff:lOff});
    p += 46+nLen+eLen+kLen;
  }
  return {
    names: entries.map(function(e){ return e.name; }),
    get: function(name){
      var e = null;
      for (var j=0;j<entries.length;j++) if (entries[j].name===name){ e=entries[j]; break; }
      if (!e) return Promise.resolve(null);
      var ln = dv.getUint16(e.lOff+26,true), le = dv.getUint16(e.lOff+28,true);
      var start = e.lOff+30+ln+le;
      var raw = u.subarray(start, start+e.cSize);
      if (e.method === 0) return Promise.resolve(raw);
      if (e.method === 8) return inflateRaw(raw);
      return Promise.reject(new Error('不支援的壓縮方式 '+e.method+'：'+name));
    }
  };
}

/* ============================================================
   2. XLSX / CSV 解析
   ============================================================ */

function colToIdx(s){
  var n=0; for (var i=0;i<s.length;i++) n = n*26 + (s.charCodeAt(i)-64);
  return n-1;
}

function unesc(s){
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"')
          .replace(/&apos;/g,"'")
          .replace(/&#x([0-9a-fA-F]+);/g,function(_,h){return String.fromCodePoint(parseInt(h,16));})
          .replace(/&#(\d+);/g,function(_,d){return String.fromCodePoint(+d);})
          .replace(/&amp;/g,'&');
}

function parseSharedStrings(xml){
  var out=[], re=/<si\b[^>]*>([\s\S]*?)<\/si>/g, m;
  while ((m = re.exec(xml))){
    var txt='', tre=/<t\b[^>]*>([\s\S]*?)<\/t>/g, tm;
    while ((tm = tre.exec(m[1]))) txt += unesc(tm[1]);
    out.push(txt);
  }
  return out;
}

function parseSheet(xml, sst){
  var grid = [], maxCol = 0;
  var re = /<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, m;
  while ((m = re.exec(xml))){
    var col = colToIdx(m[1]), row = (+m[2]) - 1, attrs = m[3]||'', body = m[4];
    if (body === undefined || body === '') continue;
    var tm = /t="([^"]+)"/.exec(attrs), type = tm ? tm[1] : 'n';
    var val = null;
    if (type === 'inlineStr'){
      var it='', tre=/<t\b[^>]*>([\s\S]*?)<\/t>/g, t2;
      while ((t2 = tre.exec(body))) it += unesc(t2[1]);
      val = it;
    } else {
      var vm = /<v>([\s\S]*?)<\/v>/.exec(body);
      if (!vm) continue;
      var raw = unesc(vm[1]);
      if (type === 's') val = sst[+raw];
      else if (type === 'str' || type === 'e') val = raw;
      else if (type === 'b') val = raw === '1';
      else { var f = parseFloat(raw); val = isNaN(f) ? raw : f; }
    }
    if (val === null || val === '') continue;
    if (!grid[row]) grid[row] = [];
    grid[row][col] = val;
    if (col > maxCol) maxCol = col;
  }
  return gridToObjects(grid, maxCol);
}

function gridToObjects(grid, maxCol){
  var hdrRow = -1, r;
  for (r=0;r<grid.length;r++){ if (grid[r] && grid[r].length){ hdrRow = r; break; } }
  if (hdrRow < 0) return [];
  var head = [];
  for (var c=0;c<=maxCol;c++){
    var h = grid[hdrRow][c];
    head[c] = (h === undefined || h === null) ? '' : String(h).trim();
  }
  var out = [];
  for (r = hdrRow+1; r < grid.length; r++){
    var row = grid[r];
    if (!row) continue;
    var o = {}, any = false;
    for (var k=0;k<=maxCol;k++){
      if (!head[k]) continue;
      var v = row[k];
      if (v === undefined) continue;
      o[head[k]] = v; any = true;
    }
    if (any) out.push(o);
  }
  return out;
}

function readXlsx(buf){
  var zip = zipRead(buf), sst = [];
  return zip.get('xl/sharedStrings.xml')
    .then(function(b){ if (b) sst = parseSharedStrings(TD.decode(b)); })
    .then(function(){ return zip.get('xl/workbook.xml'); })
    .then(function(wbB){
      var target = 'xl/worksheets/sheet1.xml';
      if (!wbB) return target;
      return zip.get('xl/_rels/workbook.xml.rels').then(function(relB){
        var sm = /<sheet\b[^>]*?r:id="([^"]+)"[^>]*\/?>/.exec(TD.decode(wbB));
        if (!sm || !relB) return target;
        var rre = new RegExp('Id="'+sm[1]+'"[^>]*Target="([^"]+)"');
        var rm = rre.exec(TD.decode(relB));
        if (!rm) {
          rre = new RegExp('Target="([^"]+)"[^>]*Id="'+sm[1]+'"');
          rm = rre.exec(TD.decode(relB));
        }
        if (!rm) return target;
        var t = rm[1].replace(/^\//,'');
        return t.indexOf('xl/') === 0 ? t : 'xl/'+t;
      });
    })
    .then(function(path){
      return zip.get(path).then(function(b){
        if (!b) throw new Error('找不到工作表 '+path);
        return parseSheet(TD.decode(b), sst);
      });
    });
}

function readCsv(text){
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  var rows = [], cur = [''], q = false, i;
  for (i=0;i<text.length;i++){
    var ch = text[i];
    if (q){
      if (ch === '"'){ if (text[i+1] === '"'){ cur[cur.length-1]+='"'; i++; } else q = false; }
      else cur[cur.length-1] += ch;
    } else if (ch === '"'){ q = true; }
    else if (ch === ','){ cur.push(''); }
    else if (ch === '\n'){ rows.push(cur); cur = ['']; }
    else if (ch !== '\r'){ cur[cur.length-1] += ch; }
  }
  if (cur.length > 1 || cur[0] !== '') rows.push(cur);
  if (!rows.length) return [];
  var head = rows[0].map(function(h){ return h.trim(); });
  return rows.slice(1).filter(function(r){ return r.some(function(v){ return v !== ''; }); })
    .map(function(r){
      var o = {};
      for (var k=0;k<head.length;k++){
        if (!head[k]) continue;
        var v = r[k];
        if (v === undefined || v === '') continue;
        var f = Number(v);
        o[head[k]] = (v.trim() !== '' && !isNaN(f)) ? f : v;
      }
      return o;
    });
}
