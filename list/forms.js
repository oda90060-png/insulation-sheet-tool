// 中間リスト → 社内書式 B02～B08
// テンプレートの1ページ分（1～H行）を盤ごとのシートに縦に複製し、H行ごとに改ページする。
// 書式のファイル構造をそのまま使うので、罫線・フォント・列幅・印刷設定はテンプレートのまま。
import { XlsxBook, XlsxSheet, copyCellFrom, clearCell, colLetter, parseAddr } from './xlsxtpl.js';

export const FORM_DEFS = {
  B02: { label: 'B02 低圧幹線絶縁抵抗測定表', match: /^B0?2[_\s-]/i, out: 'B02_低圧幹線絶縁抵抗測定表_作成.xlsx' },
  B03: { label: 'B03 低圧幹線チェックリスト', match: /^B0?3(-1)?[_\s]/i, out: 'B03_低圧幹線チェックリスト_作成.xlsx' },
  B04: { label: 'B04 動力回路絶縁抵抗測定表', match: /^B0?4[_\s-]/i, out: 'B04_動力回路絶縁抵抗測定表_作成.xlsx' },
  B05: { label: 'B05 動力回路チェックリスト', match: /^B0?5[_\s-]/i, out: 'B05_動力回路チェックリスト_作成.xlsx' },
  B06: { label: 'B06 電灯コンセント回路絶縁抵抗測定表', match: /^B0?6[_\s-]/i, out: 'B06_電灯コンセント回路絶縁抵抗測定表_作成.xlsx' },
  B07: { label: 'B07 電灯回路チェックリスト', match: /^B0?7[_\s-]/i, out: 'B07_電灯回路チェックリスト_作成.xlsx' },
  B08: { label: 'B08 コンセント回路チェックリスト', match: /^B0?8[_\s-]/i, out: 'B08_コンセント回路チェックリスト_作成.xlsx' },
};

/** ファイル名から書式コードを推定（B03-2 受変電設備更新時 は対象外） */
export function guessFormCode(name) {
  if (/^B0?3-2/i.test(name)) return null;
  for (const [k, d] of Object.entries(FORM_DEFS)) if (d.match.test(name)) return k;
  return null;
}

// ---------------------------------------------------------------- 表記
const s = v => (v === null || v === undefined) ? '' : (typeof v === 'number' && Number.isInteger(v) ? String(v) : String(v).trim());
const nfkc = t => s(t).normalize('NFKC');
const zenLen = t => [...t].reduce((n, ch) => n + (/[ᄀ-ᅟ⺀-꓏가-힣豈-﫿︰-﹏＀-｠￠-￦　-〿①-⓿■-➿]/.test(ch) ? 2 : 1), 0);

export function brkParts(t) {
  t = nfkc(t).toUpperCase();
  const typ = /ELCB|ELB|ELR/.test(t) ? 'ELB' : (/MCCB|MCB/.test(t) ? 'MCB' : '');
  // 後読み (?<!) は古い iPad の Safari で使えないため、直前の文字をグループで判定する
  const mp = t.match(/(?:^|[^\d.])(\d)\s*P(?![A-Z])/) || t.match(/(?:^|[^\d.])(\d)P/);
  const maf = t.match(/(?:^|[^\d-])(\d+)\s*AF/);            // 装置番号（ELCB-101 等）は '-' 付きなので除外
  const mat = t.match(/\/\s*(\d+)\s*A(?:T|\b)/) || t.match(/(?:^|[^\d-])(\d+)\s*AT/);
  return { typ, poles: mp ? mp[1] + 'P' : '', af: maf ? maf[1] : '', at: mat ? mat[1] : '' };
}

/** style: 'b06' → 'ELB2P 50/20AT' / 'b03' → 'MCB 3P\n225AF/200AT' */
export function brkShort(t, style) {
  const { typ, poles, af, at } = brkParts(t);
  if (!af && !at) return s(t);
  if (style === 'b06') return `${typ}${poles} ${af}/${at}AT`.trim();
  if (style === 'b03') return `${typ} ${poles}\n${af}AF/${at}AT`.trim();
  return `${typ} ${poles} ${af}AF/${at}AT`.trim();
}
const voltTxt = v => { v = s(v); return (!v || !/^[\d/]+$/.test(v)) ? v : v + 'V'; };

// ---------------------------------------------------------------- 書式の複製
class Form {
  /** base: テンプレートのシート番号、H: 1ページの行数、lastCol: 印刷範囲の最終列、header: 共通ヘッダの記入先 */
  static async open(buf, base, H, lastCol, header) {
    const f = new Form();
    f.book = await XlsxBook.open(buf);
    f.baseDoc = await f.book.sheetDoc(base);
    f.H = H; f.lastCol = lastCol; f.header = header;
    f.used = new Set();
    return f;
  }

  sheetName(name) {
    const base = (name.replace(/[\[\]:*?/\\]/g, '-').slice(0, 31)) || '盤';
    let n = base, i = 2;
    while (this.used.has(n)) n = `${base.slice(0, 28)}(${i++})`;
    this.used.add(n);
    return n;
  }

  newSheet(panel, pages) {
    const ws = new XlsxSheet(this.book, this.baseDoc, this.H, this.lastCol, pages);
    this.book.addSheet(this.sheetName(panel), ws);
    return ws;
  }

  /** 値を書く。fit=true なら枠に収まるよう折返し＋フォント縮小（9.5pt→最小6pt） */
  put(ws, addr, value, fit = false) {
    const text = (value === null || value === undefined) ? '' : String(value);
    ws.setText(addr, text);
    if (!text || (!fit && !text.includes('\n'))) return;
    const st = ws.style(addr);
    const base = this.book.fontSize(st);
    let size = base;
    if (fit) {
      const { w, h } = ws.box(addr);
      size = 6;
      for (const z of [base, 9, 8.5, 8, 7.5, 7, 6.5, 6]) {
        if (z > base) continue;
        const perLine = Math.max(1, w * 1.05 * 11 / z);
        const lines = text.split('\n').reduce((n, t) => n + Math.max(1, Math.ceil(zenLen(t) / perLine)), 0);
        if (lines * z * 1.22 <= h) { size = z; break; }
      }
    }
    ws.setStyle(addr, this.book.styleVariant(st, { size, wrap: true }));
  }

  fillCommon(ws, pages, common) {
    const era = common['年号'] || '令和';
    const maxC = parseAddr(this.lastCol + '1').c;
    for (let k = 0; k < pages; k++) {
      const off = k * this.H;
      for (let r = 1; r <= 4; r++) for (let c = 1; c <= maxC; c++) {
        const a = colLetter(c) + (r + off);
        const t = ws.cell(a, false) ? ws.text(a) : '';
        if (t.includes('平成')) ws.setText(a, t.replace('平成', era));
      }
      for (const [key, addr] of Object.entries(this.header)) {
        const v = common[key];
        if (v) { const { r, c } = parseAddr(addr); this.put(ws, colLetter(c) + (r + off), v); }
      }
    }
  }

  blob() { return this.book.toBlob(); }
}

const pagesFor = (n, per) => Math.max(1, Math.ceil(n / per));

function groupByPanel(rows, key, order) {
  const idx = new Map(order.map((p, i) => [p, i]));
  const groups = new Map();
  rows.forEach((r, i) => {
    const p = s(r[key]);
    if (!p) return;
    if (!groups.has(p)) groups.set(p, []);
    const n = parseFloat(r['並び']);
    groups.get(p).push({ k: Number.isFinite(n) ? n : 1e9, i, r });
  });
  const names = [...groups.keys()];
  names.sort((a, b) => ((idx.has(a) ? idx.get(a) : 1e6) - (idx.has(b) ? idx.get(b) : 1e6)) || (names.indexOf(a) - names.indexOf(b)));
  return names.map(p => [p, groups.get(p).sort((a, b) => a.k - b.k || a.i - b.i).map(x => x.r)]);
}

// ---------------------------------------------------------------- 各書式
async function makeB02(buf, mid, order, common) {
  const rows = mid.trunk.filter(r => !['予備', '除外'].includes(s(r['区分'])));
  const f = await Form.open(buf, 0, 27, 'BE', { '立会者': 'X1', '実施者': 'X2' });
  for (const [p, items] of groupByPanel(rows, '配電盤', order)) {
    const n = pagesFor(items.length, 8);
    const ws = f.newSheet(p, n);
    f.fillCommon(ws, n, common);
    items.forEach((r, i) => {
      const row = Math.floor(i / 8) * 27 + 5 + 2 * (i % 8);
      const a = [s(r['幹線番号']), (s(r['電気方式']) + ' ' + voltTxt(r['電圧'])).trim(), s(r['幹線サイズ'])].filter(Boolean).join('\n');
      f.put(ws, `A${row}`, a, true);
      const rng = s(r['測定範囲']) || `${p} ～ ${s(r['行先'])}`;
      f.put(ws, `I${row}`, rng.includes('\n') ? rng : rng.replace('～ ', '～').replace('～', '～\n'), true);
    });
  }
  return f;
}

async function makeB03(buf, mid, order, common) {
  const rows = mid.trunk.filter(r => !['予備', '除外'].includes(s(r['区分'])));
  const f = await Form.open(buf, 0, 23, 'BE', { '立会者': 'AC1', '実施者': 'AS1' });
  for (const [p, items] of groupByPanel(rows, '配電盤', order)) {
    const n = pagesFor(items.length, 10);
    const ws = f.newSheet(p, n);
    f.fillCommon(ws, n, common);
    items.forEach((r, i) => {
      const row = Math.floor(i / 10) * 23 + 7 + (i % 10);
      const allow = s(r['許容電流(A)']);
      const a = [s(r['幹線番号']), (s(r['電気方式']) + ' ' + voltTxt(r['電圧'])).trim(), s(r['幹線サイズ']), allow ? `(${allow} A)` : '(　　 A)'].filter(Boolean).join('\n');
      f.put(ws, `A${row}`, a, true);
      f.put(ws, `E${row}`, brkShort(r['ブレーカ'], 'b03'), true);
      f.put(ws, `K${row}`, s(r['測定範囲']) || `${p} ～ ${s(r['行先'])}`, true);
    });
  }
  return f;
}

async function makeB04(buf, mid, order, common) {
  const rows = mid.power.filter(r => s(r['区分']) !== '除外');
  const f = await Form.open(buf, 0, 28, 'BE', { '立会者': 'AA1', '実施者': 'AA2' });
  for (const [p, items] of groupByPanel(rows, '盤名', order)) {
    const n = pagesFor(items.length, 8);
    const ws = f.newSheet(p, n);
    f.fillCommon(ws, n, common);
    for (let k = 0; k < n; k++) f.put(ws, `S${2 + k * 28}`, p, true);
    items.forEach((r, i) => {
      const row = Math.floor(i / 8) * 28 + 5 + 2 * (i % 8);
      f.put(ws, `A${row}`, s(r['回路番号']));
      f.put(ws, `D${row}`, [s(r['負荷名称']), s(r['負荷記号'])].filter(Boolean).join(' '), true);
      f.put(ws, `O${row}`, s(r['ケーブル']), true);
      f.put(ws, `R${row}`, s(r['電圧(V)']));
      f.put(ws, `U${row}`, s(r['容量(kW)']), true);
    });
  }
  return f;
}

async function makeB05(buf, mid, order, common) {
  const rows = mid.power.filter(r => !['予備', '除外'].includes(s(r['区分'])));
  const f = await Form.open(buf, 0, 24, 'BL', { '立会者': 'AS1', '実施者': 'BE1' });
  for (const [p, items] of groupByPanel(rows, '盤名', order)) {
    const n = pagesFor(items.length, 10);
    const ws = f.newSheet(p, n);
    f.fillCommon(ws, n, common);
    for (let k = 0; k < n; k++) f.put(ws, `AH${2 + k * 24}`, p, true);
    items.forEach((r, i) => {
      const row = Math.floor(i / 10) * 24 + 7 + (i % 10);
      f.put(ws, `A${row}`, s(r['回路番号']));
      f.put(ws, `C${row}`, [s(r['負荷名称']), s(r['負荷記号'])].filter(Boolean).join('\n'), true);
      const allow = s(r['許容電流(A)']), cab = s(r['ケーブル']);
      f.put(ws, `G${row}`, cab ? cab + (allow ? `\n(${allow}A)` : '\n(　 A)') : '', true);
      f.put(ws, `J${row}`, brkShort(r['ブレーカ'], 'b03'), true);
      f.put(ws, `M${row}`, s(r['電圧(V)']));
      f.put(ws, `O${row}`, s(r['容量(kW)']), true);
      f.put(ws, `Q${row}`, s(r['負荷電流(A)']), true);
    });
  }
  return f;
}

/** 左＝奇数・右＝偶数になるよう、回路記号（◎・○・□F…）が変わる所で左列から始める */
function pairLayout(items) {
  const out = []; let prev = null;
  for (const r of items) {
    const no = s(r['回路番号']);
    const g = /\d/.test(no) ? no.replace(/\d.*$/, '') : no.replace(/[A-Za-z]+$/, '*');
    if (prev !== null && g !== prev && out.length % 2 === 1) out.push(null);
    out.push(r); prev = g;
  }
  return out;
}

async function makeB06(buf, mid, order, common) {
  const rows = mid.light.filter(r => s(r['区分']) !== '除外');
  const pinfo = new Map(mid.panels.map(p => [s(p['盤名']), p]));
  const f = await Form.open(buf, 0, 37, 'BF', { '立会者': 'Y1', '実施者': 'Y2' });
  // テンプレートに残っている前回の記入値を消し、記入例シートから「Ｂ６」「幹線番号」を補う
  const base = new XlsxSheet(f.book, f.baseDoc, 37, 'BF', 1);
  for (const a of ['Y1', 'Y2', 'AW2', 'AZ2', 'BD2', 'AJ3', 'Y3']) clearCell(base, a);
  base.setText('AK2', '　令和　　年　　月　　日');
  for (let r = 7; r <= 30; r++) for (let c = 1; c <= 58; c++) clearCell(base, colLetter(c) + r);
  if (f.book.sheetInfo.length > 2) {
    const ex = await f.book.sheetDoc(2);
    for (const a of ['A1', 'V3']) copyCellFrom(ex, a, base);
  }
  f.baseDoc = base.doc;
  for (const [p, itemsRaw] of groupByPanel(rows, '盤名', order)) {
    const items = pairLayout(itemsRaw);
    const n = pagesFor(items.length, 24);
    const ws = f.newSheet(p, n);
    const info = pinfo.get(p) || {};
    f.fillCommon(ws, n, common);
    for (let k = 0; k < n; k++) {
      const off = k * 37;
      f.put(ws, `B${3 + off}`, `盤名称：　${p}`);
      f.put(ws, `Y${3 + off}`, s(info['電源(幹線番号)']), true);
      f.put(ws, `AJ${3 + off}`, s(info['幹線サイズ']), true);
    }
    items.forEach((r, i) => {
      if (!r) return;
      const pg = Math.floor(i / 24), j = i % 24;
      const row = pg * 37 + 7 + 2 * Math.floor(j / 2);
      const name = s(r['負荷名称']);
      if (j % 2 === 0) {
        f.put(ws, `Q${row}`, name, true);
        f.put(ws, `V${row}`, brkShort(r['ブレーカ'], 'b06'), true);
        f.put(ws, `Y${row}`, s(r['回路番号']), true);
        f.put(ws, `AA${row}`, voltTxt(r['電圧(V)']));
      } else {
        f.put(ws, `AE${row}`, voltTxt(r['電圧(V)']));
        f.put(ws, `AG${row}`, s(r['回路番号']), true);
        f.put(ws, `AI${row}`, brkShort(r['ブレーカ'], 'b06'), true);
        f.put(ws, `AL${row}`, name, true);
      }
    });
  }
  return f;
}

async function makeB07B08(buf, mid, order, common, kind) {
  const conf = kind === '電灯'
    ? { H: 25, first: 5, per: 14, last: 'BC', no: 'A', name: 'E', v: 'T', panel: 'Z2', hdr: { '立会者': 'Z1' } }
    : { H: 27, first: 7, per: 15, last: 'BA', no: 'A', name: 'C', v: 'P', panel: 'AB2', hdr: { '立会者': 'AB1' } };
  const rows = mid.light.filter(r => s(r['区分']) === kind);
  const f = await Form.open(buf, 0, conf.H, conf.last, conf.hdr);
  const pc = parseAddr(conf.panel);
  for (const [p, items] of groupByPanel(rows, '盤名', order)) {
    const n = pagesFor(items.length, conf.per);
    const ws = f.newSheet(p, n);
    f.fillCommon(ws, n, common);
    for (let k = 0; k < n; k++) f.put(ws, colLetter(pc.c) + (pc.r + k * conf.H), p);
    items.forEach((r, i) => {
      const row = Math.floor(i / conf.per) * conf.H + conf.first + (i % conf.per);
      f.put(ws, `${conf.no}${row}`, s(r['回路番号']), true);
      f.put(ws, `${conf.name}${row}`, s(r['負荷名称']), true);
      f.put(ws, `${conf.v}${row}`, s(r['電圧(V)']));
    });
  }
  return f;
}

const MAKERS = {
  B02: makeB02, B03: makeB03, B04: makeB04, B05: makeB05, B06: makeB06,
  B07: (b, m, o, c) => makeB07B08(b, m, o, c, '電灯'),
  B08: (b, m, o, c) => makeB07B08(b, m, o, c, 'コンセント'),
};

/** code: 'B02'…'B08', templateBuf: ArrayBuffer, mid: {panels,trunk,power,light}, common: {立会者,実施者,年号} → Blob */
export async function makeForm(code, templateBuf, mid, common) {
  const order = mid.panels.map(p => s(p['盤名']));
  const f = await MAKERS[code](templateBuf.slice(0), mid, order, common || {});
  if (!f.book.out.length) return null;
  return f.blob();
}
