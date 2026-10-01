// 社内書式の xlsx を「中身をそのまま複製して値だけ書き込む」ための小さなライブラリ（JSZip + DOM）。
// ExcelJS で読み書きすると、標準フォントの置き換え（列幅が変わりA4からはみ出す）、
// 結合セルの罫線の消失、9.5pt などの小数フォントの丸めが起きるため、XML を直接扱う。
/* global JSZip */

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const T_WS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const parse = t => new DOMParser().parseFromString(t, 'application/xml');
const ser = d => XML_HEAD + new XMLSerializer().serializeToString(d).replace(/^<\?xml[^>]*\?>\s*/, '');
const kids = (el, name) => Array.from(el.childNodes).filter(n => n.nodeType === 1 && n.localName === name);
const kid = (el, name) => kids(el, name)[0] || null;

export function colIndex(letters) { let n = 0; for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
export function colLetter(n) { let r = ''; while (n > 0) { const m = (n - 1) % 26; r = String.fromCharCode(65 + m) + r; n = Math.floor((n - 1) / 26); } return r; }
export function parseAddr(a) { const m = /^([A-Z]+)(\d+)$/.exec(a.replace(/\$/g, '')); return { c: colIndex(m[1]), r: +m[2] }; }
export function parseRange(rg) { const [a, b] = rg.split(':'); const p = parseAddr(a), q = parseAddr(b || a); return { r1: p.r, c1: p.c, r2: q.r, c2: q.c }; }

export class XlsxBook {
  static async open(buf) {
    const b = new XlsxBook();
    b.zip = await JSZip.loadAsync(buf);
    b.wb = parse(await b.zip.file('xl/workbook.xml').async('string'));
    b.rels = parse(await b.zip.file('xl/_rels/workbook.xml.rels').async('string'));
    b.ct = parse(await b.zip.file('[Content_Types].xml').async('string'));
    b.styles = parse(await b.zip.file('xl/styles.xml').async('string'));
    const sstFile = b.zip.file('xl/sharedStrings.xml');
    b.sst = [];
    if (sstFile) {
      const sd = parse(await sstFile.async('string'));
      for (const si of sd.getElementsByTagNameNS(NS, 'si')) b.sst.push(Array.from(si.getElementsByTagNameNS(NS, 't')).map(t => t.textContent).join(''));
    }
    const relById = {};
    for (const r of b.rels.getElementsByTagNameNS(NS_PKG, 'Relationship')) relById[r.getAttribute('Id')] = r.getAttribute('Target');
    b.sheetInfo = Array.from(b.wb.getElementsByTagNameNS(NS, 'sheet')).map(s => {
      const t = relById[s.getAttributeNS(NS_R, 'id')];
      return { name: s.getAttribute('name'), path: t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '') };
    });
    b.fontsEl = b.styles.getElementsByTagNameNS(NS, 'fonts')[0];
    b.xfsEl = b.styles.getElementsByTagNameNS(NS, 'cellXfs')[0];
    b.xfCache = new Map();
    b.out = [];
    return b;
  }

  async sheetDoc(i) { return parse(await this.zip.file(this.sheetInfo[i].path).async('string')); }

  fontSize(s) {
    const xf = kids(this.xfsEl, 'xf')[s || 0];
    const font = kids(this.fontsEl, 'font')[+(xf && xf.getAttribute('fontId')) || 0];
    const sz = font && kid(font, 'sz');
    return sz ? parseFloat(sz.getAttribute('val')) : 11;
  }

  /** 既存スタイル s を元に、フォントサイズ・折返しを変えたスタイル番号を返す */
  styleVariant(s, { size, wrap }) {
    const key = `${s}|${size}|${wrap}`;
    if (this.xfCache.has(key)) return this.xfCache.get(key);
    const xfs = kids(this.xfsEl, 'xf');
    const xf = xfs[s || 0].cloneNode(true);
    if (size && size !== this.fontSize(s)) {
      const fonts = kids(this.fontsEl, 'font');
      const f = fonts[+xf.getAttribute('fontId') || 0].cloneNode(true);
      let sz = kid(f, 'sz');
      if (!sz) { sz = this.styles.createElementNS(NS, 'sz'); f.insertBefore(sz, f.firstChild); }
      sz.setAttribute('val', String(size));
      this.fontsEl.appendChild(f);
      xf.setAttribute('fontId', String(fonts.length));
      xf.setAttribute('applyFont', '1');
    }
    if (wrap) {
      let al = kid(xf, 'alignment');
      if (!al) { al = this.styles.createElementNS(NS, 'alignment'); xf.insertBefore(al, xf.firstChild); }
      if (!al.getAttribute('horizontal')) al.setAttribute('horizontal', 'center');
      if (!al.getAttribute('vertical')) al.setAttribute('vertical', 'center');
      al.setAttribute('wrapText', '1');
      xf.setAttribute('applyAlignment', '1');
    }
    this.xfsEl.appendChild(xf);
    const idx = xfs.length;
    this.xfCache.set(key, idx);
    return idx;
  }

  addSheet(name, sheet) { this.out.push({ name, sheet }); }

  async toBlob() {
    const z = this.zip;
    // 元のシート・計算チェーン・プリンター設定を取り除き、新しいシートに差し替える
    for (const p of Object.keys(z.files)) {
      if (/^xl\/worksheets\//.test(p) || p === 'xl/calcChain.xml' || /^xl\/printerSettings\//.test(p)) z.remove(p);
    }
    const relRoot = this.rels.documentElement;
    for (const r of Array.from(relRoot.getElementsByTagNameNS(NS_PKG, 'Relationship'))) {
      if (r.getAttribute('Type') === T_WS || /calcChain$/.test(r.getAttribute('Type'))) relRoot.removeChild(r);
    }
    const ctRoot = this.ct.documentElement;
    for (const o of Array.from(ctRoot.getElementsByTagNameNS(NS_CT, 'Override'))) {
      const pn = o.getAttribute('PartName');
      if (/^\/xl\/worksheets\//.test(pn) || pn === '/xl/calcChain.xml') ctRoot.removeChild(o);
    }
    const wbRoot = this.wb.documentElement;
    const sheetsEl = wbRoot.getElementsByTagNameNS(NS, 'sheets')[0];
    while (sheetsEl.firstChild) sheetsEl.removeChild(sheetsEl.firstChild);
    let dn = wbRoot.getElementsByTagNameNS(NS, 'definedNames')[0];
    if (dn) wbRoot.removeChild(dn);
    dn = this.wb.createElementNS(NS, 'definedNames');
    this.out.forEach(({ name, sheet }, i) => {
      const n = i + 1, rid = `rIdSheet${n}`;
      const se = this.wb.createElementNS(NS, 'sheet');
      se.setAttribute('name', name); se.setAttribute('sheetId', String(n)); se.setAttributeNS(NS_R, 'r:id', rid);
      sheetsEl.appendChild(se);
      const rel = this.rels.createElementNS(NS_PKG, 'Relationship');
      rel.setAttribute('Id', rid); rel.setAttribute('Type', T_WS); rel.setAttribute('Target', `worksheets/sheet${n}.xml`);
      relRoot.appendChild(rel);
      const ov = this.ct.createElementNS(NS_CT, 'Override');
      ov.setAttribute('PartName', `/xl/worksheets/sheet${n}.xml`);
      ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
      ctRoot.appendChild(ov);
      if (sheet.printArea) {
        const d = this.wb.createElementNS(NS, 'definedName');
        d.setAttribute('name', '_xlnm.Print_Area'); d.setAttribute('localSheetId', String(i));
        d.textContent = `'${name.replace(/'/g, "''")}'!${sheet.printArea}`;
        dn.appendChild(d);
      }
      sheet.finalize(i === 0);
      z.file(`xl/worksheets/sheet${n}.xml`, ser(sheet.doc));
    });
    if (dn.childNodes.length) {
      // definedNames は sheets の直後（calcPr より前）に置く
      const after = sheetsEl.nextSibling;
      let ref = after;
      while (ref && ref.nodeType === 1 && ['functionGroups', 'externalReferences'].includes(ref.localName)) ref = ref.nextSibling;
      wbRoot.insertBefore(dn, ref);
    }
    const bv = wbRoot.getElementsByTagNameNS(NS, 'workbookView')[0];
    if (bv) { bv.setAttribute('activeTab', '0'); bv.removeAttribute('firstSheet'); }
    for (const [el, name] of [[this.fontsEl, 'font'], [this.xfsEl, 'xf']]) el.setAttribute('count', String(kids(el, name).length));
    z.file('xl/workbook.xml', ser(this.wb));
    z.file('xl/_rels/workbook.xml.rels', ser(this.rels));
    z.file('[Content_Types].xml', ser(this.ct));
    z.file('xl/styles.xml', ser(this.styles));
    if (z.file('docProps/app.xml')) {
      z.file('docProps/app.xml', XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application></Properties>');
    }
    return z.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
  }
}

/** テンプレートの1ページ分（1～H行）を、指定ページ数だけ縦に複製したシート */
export class XlsxSheet {
  constructor(book, baseDoc, H, lastCol, pages) {
    this.book = book; this.H = H; this.lastCol = lastCol; this.pages = pages;
    const doc = baseDoc.cloneNode(true);
    this.doc = doc;
    const root = doc.documentElement;
    const sd = root.getElementsByTagNameNS(NS, 'sheetData')[0];
    const baseRows = kids(sd, 'row').filter(r => +r.getAttribute('r') <= H);
    while (sd.firstChild) sd.removeChild(sd.firstChild);
    this.rows = new Map(); this.cells = new Map();
    for (let k = 0; k < pages; k++) {
      const off = k * H;
      for (const br of baseRows) {
        const row = br.cloneNode(true);
        const r = +br.getAttribute('r') + off;
        row.setAttribute('r', String(r));
        for (const c of kids(row, 'c')) {
          const a = parseAddr(c.getAttribute('r'));
          const addr = colLetter(a.c) + r;
          c.setAttribute('r', addr);
          this.cells.set(addr, c);
        }
        sd.appendChild(row); this.rows.set(r, row);
      }
    }
    this.sd = sd;
    // 結合セル（1ページ分）を各ページに複製
    const mcOld = root.getElementsByTagNameNS(NS, 'mergeCells')[0];
    this.baseMerges = mcOld ? Array.from(mcOld.getElementsByTagNameNS(NS, 'mergeCell')).map(m => parseRange(m.getAttribute('ref'))).filter(m => m.r2 <= H) : [];
    const mc = doc.createElementNS(NS, 'mergeCells');
    for (let k = 0; k < pages; k++) for (const m of this.baseMerges) {
      const e = doc.createElementNS(NS, 'mergeCell');
      e.setAttribute('ref', `${colLetter(m.c1)}${m.r1 + k * H}:${colLetter(m.c2)}${m.r2 + k * H}`);
      mc.appendChild(e);
    }
    mc.setAttribute('count', String(mc.childNodes.length));
    if (mcOld) root.replaceChild(mc, mcOld);
    else if (mc.childNodes.length) root.insertBefore(mc, sd.nextSibling);
    // 列幅・行高（文字の収まり計算用）
    const fmt = root.getElementsByTagNameNS(NS, 'sheetFormatPr')[0];
    this.defColW = fmt && fmt.getAttribute('defaultColWidth') ? parseFloat(fmt.getAttribute('defaultColWidth')) : 8.43;
    this.defRowH = fmt && fmt.getAttribute('defaultRowHeight') ? parseFloat(fmt.getAttribute('defaultRowHeight')) : 13.5;
    this.colW = new Map();
    for (const c of root.getElementsByTagNameNS(NS, 'col')) {
      const w = parseFloat(c.getAttribute('width'));
      for (let i = +c.getAttribute('min'); i <= Math.min(+c.getAttribute('max'), 400); i++) this.colW.set(i, w);
    }
    this.printArea = `$A$1:$${lastCol}$${pages * H}`;
  }

  /**
   * 1ページ分（H行×印刷列）が A4 横に収まる倍率。書式の倍率より小さいときだけ下げる。
   * 列幅→ピクセルは Excel の計算式で、標準フォントの数字幅を 8px（ＭＳ Ｐゴシック 11pt の大きい方）として安全側に見積もる。
   */
  fitScale() {
    const root = this.doc.documentElement;
    const pm = root.getElementsByTagNameNS(NS, 'pageMargins')[0];
    const mg = k => (pm && pm.getAttribute(k) !== null ? parseFloat(pm.getAttribute(k)) : 0.75) * 72;
    const ps = root.getElementsByTagNameNS(NS, 'pageSetup')[0];
    const cur = ps && ps.getAttribute('scale') ? parseInt(ps.getAttribute('scale'), 10) : 100;
    const land = !ps || ps.getAttribute('orientation') !== 'portrait';
    const [pw0, ph0] = land ? [841.89, 595.28] : [595.28, 841.89];
    const pw = pw0 - mg('left') - mg('right'), ph = ph0 - mg('top') - mg('bottom');
    let hpt = 0;
    for (let r = 1; r <= this.H; r++) { const row = this.rows.get(r); const ht = row && row.getAttribute('ht'); hpt += ht ? parseFloat(ht) : this.defRowH; }
    let px = 0;
    const last = parseAddr(this.lastCol + '1').c;
    for (let c = 1; c <= last; c++) px += Math.trunc(((256 * (this.colW.get(c) || this.defColW) + Math.trunc(128 / 8)) / 256) * 8);
    const fit = Math.floor(100 * Math.min(ph / hpt, pw / (px * 0.75)));
    return Math.max(10, Math.min(cur, fit));
  }

  /** シートを仕上げる（改ページ・範囲・選択状態・倍率） */
  finalize(first) {
    const doc = this.doc, root = doc.documentElement;
    const ps0 = root.getElementsByTagNameNS(NS, 'pageSetup')[0];
    if (ps0) ps0.setAttribute('scale', String(this.fitScale()));
    const dim = root.getElementsByTagNameNS(NS, 'dimension')[0];
    if (dim) dim.setAttribute('ref', `A1:${this.lastCol}${this.pages * this.H}`);
    for (const v of root.getElementsByTagNameNS(NS, 'sheetView')) {
      if (first) v.setAttribute('tabSelected', '1'); else v.removeAttribute('tabSelected');
    }
    const ps = root.getElementsByTagNameNS(NS, 'pageSetup')[0];
    if (ps) ps.removeAttributeNS(NS_R, 'id');
    for (const name of ['rowBreaks', 'drawing', 'legacyDrawing']) { const e = root.getElementsByTagNameNS(NS, name)[0]; if (e) root.removeChild(e); }
    if (this.pages > 1) {
      const rb = doc.createElementNS(NS, 'rowBreaks');
      for (let k = 1; k < this.pages; k++) {
        const b = doc.createElementNS(NS, 'brk');
        b.setAttribute('id', String(k * this.H)); b.setAttribute('max', '16383'); b.setAttribute('man', '1');
        rb.appendChild(b);
      }
      rb.setAttribute('count', String(this.pages - 1)); rb.setAttribute('manualBreakCount', String(this.pages - 1));
      // 要素の並び順（headerFooter の後、colBreaks・drawing 等の前）
      const order = ['colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
      let ref = null;
      for (const n of Array.from(root.childNodes)) if (n.nodeType === 1 && order.includes(n.localName)) { ref = n; break; }
      root.insertBefore(rb, ref);
    }
  }

  cell(addr, create = true) {
    let c = this.cells.get(addr);
    if (c || !create) return c;
    const { r, c: col } = parseAddr(addr);
    let row = this.rows.get(r);
    if (!row) {
      row = this.doc.createElementNS(NS, 'row'); row.setAttribute('r', String(r));
      const next = Array.from(this.rows.keys()).filter(k => k > r).sort((a, b) => a - b)[0];
      this.sd.insertBefore(row, next ? this.rows.get(next) : null);
      this.rows.set(r, row);
    }
    c = this.doc.createElementNS(NS, 'c'); c.setAttribute('r', addr);
    const after = kids(row, 'c').find(x => parseAddr(x.getAttribute('r')).c > col);
    row.insertBefore(c, after || null);
    this.cells.set(addr, c);
    return c;
  }

  text(addr) {
    const c = this.cell(addr, false);
    if (!c) return '';
    const t = c.getAttribute('t'), v = kid(c, 'v');
    if (t === 's') return v ? (this.book.sst[+v.textContent] || '') : '';
    if (t === 'inlineStr') { const is = kid(c, 'is'); return is ? Array.from(is.getElementsByTagNameNS(NS, 't')).map(x => x.textContent).join('') : ''; }
    return v ? v.textContent : '';
  }

  setText(addr, value) {
    const c = this.cell(addr);
    for (const n of Array.from(c.childNodes)) c.removeChild(n);
    c.removeAttribute('t');
    if (value === null || value === undefined || value === '') return c;
    c.setAttribute('t', 'inlineStr');
    const is = this.doc.createElementNS(NS, 'is'), t = this.doc.createElementNS(NS, 't');
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = String(value);
    is.appendChild(t); c.appendChild(is);
    return c;
  }

  style(addr) { const c = this.cell(addr, false); return c ? +(c.getAttribute('s') || 0) : 0; }
  setStyle(addr, s) { this.cell(addr).setAttribute('s', String(s)); }

  /** addr を含む結合範囲の幅（文字単位）と高さ（pt）。ページ位置に関係なくテンプレートの1ページ目で計算 */
  box(addr) {
    const { r, c } = parseAddr(addr);
    const rr = ((r - 1) % this.H) + 1;
    const m = this.baseMerges.find(m => rr >= m.r1 && rr <= m.r2 && c >= m.c1 && c <= m.c2);
    const c1 = m ? m.c1 : c, c2 = m ? m.c2 : c, r1 = m ? m.r1 : rr, r2 = m ? m.r2 : rr;
    let w = 0, h = 0;
    for (let i = c1; i <= c2; i++) w += this.colW.get(i) || this.defColW;
    for (let i = r1; i <= r2; i++) { const row = this.rows.get(i); const ht = row && row.getAttribute('ht'); h += ht ? parseFloat(ht) : this.defRowH; }
    return { w, h };
  }
}

/** 他シートのセル（値とスタイル）をそのまま写す（同じブック内なので共有文字列の番号も有効） */
export function copyCellFrom(srcDoc, addr, dstSheet) {
  const src = Array.from(srcDoc.getElementsByTagNameNS(NS, 'c')).find(c => c.getAttribute('r') === addr);
  const dst = dstSheet.cell(addr);
  for (const n of Array.from(dst.childNodes)) dst.removeChild(n);
  for (const a of ['t', 's']) dst.removeAttribute(a);
  if (!src) return;
  for (const a of ['t', 's']) if (src.getAttribute(a) !== null) dst.setAttribute(a, src.getAttribute(a));
  for (const n of Array.from(src.childNodes)) dst.appendChild(dstSheet.doc.importNode(n, true));
}

/** 1ページ目のセルの値を消す（スタイルは残す） */
export function clearCell(sheetOrDoc, addr) {
  const c = sheetOrDoc.cell(addr, false);
  if (!c) return;
  for (const n of Array.from(c.childNodes)) c.removeChild(n);
  c.removeAttribute('t');
}
