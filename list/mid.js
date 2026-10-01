// 読み取り結果（ページごとのJSON）→ 中間リスト、および 中間リスト.xlsx の読み書き
/* global ExcelJS */

export const COLS = {
  panels: ['盤名', '種別', '電源(幹線番号)', '電気方式', '主幹', '幹線サイズ', '出典', '備考'],
  trunk: ['並び', '配電盤', '幹線番号', '行先', '電気方式', '電圧', '幹線サイズ', '許容電流(A)', 'ブレーカ', '容量', '測定範囲', '区分', '備考', '出典'],
  power: ['並び', '盤名', '回路番号', '負荷名称', '負荷記号', '容量(kW)', '電圧(V)', 'ケーブル', '許容電流(A)', 'ブレーカ', '負荷電流(A)', '区分', '備考', '出典'],
  light: ['並び', '盤名', '回路番号', '負荷名称', '容量(VA)', '電圧(V)', 'ブレーカ', '区分', '備考', '出典'],
  checks: ['盤名', 'ページ', '内容'],
};
export const SHEETS = { panels: '盤一覧', trunk: '低圧幹線', power: '動力回路', light: '電灯コンセント', checks: '要確認' };
export const CATS = { trunk: ['幹線', '予備', '除外'], power: ['動力', '予備', '除外'], light: ['電灯', 'コンセント', 'その他', '予備', '除外'] };

const s = v => (v === null || v === undefined) ? '' : String(v).trim();
const nz = v => s(v).normalize('NFKC').trim();

export function emptyMid() { return { panels: [], trunk: [], power: [], light: [], checks: [] }; }

// ---------------------------------------------------------------- 分類
const EQUIP = /トランス|ファン|ポンプ|空調|PAC|室内機|室外機|自動ドア|盤|シャッター|ヒーター|換気|給湯|温水器|ユニット|CAV|VAV|装置|機器|制御|監視|バルブ|エアシャワー|冷蔵|冷凍|EV|エレベ|ITV|インターホン|カメラ|サーバ|UPS|充電|AMR|コンベア|ハンガー|乾燥|洗浄機|複合機|RS-|SB|PLC|端子/;

function classifyLight(name, kind) {
  const n = nz(name);
  if (kind === 'space') return '除外';
  if (n === '予備' || kind === '予備') return '予備';
  if (!n) return 'その他';
  if (n.includes('コンセント') || /\bOA\b/.test(n)) return 'コンセント';
  if (/電灯|照明|誘導灯|外灯|灯$/.test(n)) return '電灯';
  return 'その他';
}

function devNo(c) {
  const m = nz(c.breaker).toUpperCase().match(/(?:ELCB|MCCB|ELB|MCB)-?\s*([A-Z]?\d+)/);
  return m ? `(${m[1]})` : '';
}
function circNo(c) {
  if (!nz(c.no) && !nz(c.sub)) return devNo(c);
  return `${c.mark || ''}${nz(c.no)}${nz(c.sub)}`;
}
const grpOf = c => `${c.mark || ''}|${nz(c.no).replace(/\d+.*$/, '')}`;
function sortCircuits(items) {
  const groups = [];
  for (const c of items) { const g = grpOf(c); if (!groups.includes(g)) groups.push(g); }
  const num = c => { const m = nz(c.no).match(/(\d+)/); return m ? +m[1] : 9999; };
  return [...items].sort((a, b) => (groups.indexOf(grpOf(a)) - groups.indexOf(grpOf(b))) || (num(a) - num(b)) || nz(a.sub).localeCompare(nz(b.sub)) || (a._seq - b._seq));
}

function mergeInc(a, b) {
  if (!a) return { ...b };
  const out = { ...a };
  for (const k of ['source', 'cable', 'main_breaker', 'system']) {
    const va = nz(a[k]), vb = nz(b[k]);
    if (vb && !va.split(' / ').includes(vb)) out[k] = va ? `${va} / ${vb}` : vb;
  }
  return out;
}

function nominal(system) {
  // '3φ3W6450V/210V' '1φ3W 210-105V' → ['3φ3W','200'] 等
  const t = nz(system).replace(/\s/g, '');
  const ph = (t.match(/(\dφ\dW)/) || [])[1] || '';
  if (!ph) return ['', ''];
  if (ph === '1φ3W') return [ph, '200/100'];
  if (ph === '3φ4W') return [ph, /440/.test(t) ? '440' : '400'];
  const m = t.match(/\/(\d+)V/) || t.match(/(\d{3})V/);
  const v = m ? m[1] : '';
  return [ph, ({ 210: '200', 105: '100' })[v] || v];
}

/**
 * pages: [{pdfName, page, kind:'分電盤'|'動力盤'|'CUB', panel, result}]（result は AI の JSON）
 * 戻り値: 中間リスト
 */
export function buildMid(pages) {
  const mid = emptyMid();
  const src = p => `${p.pdfName} p${p.page}`;
  const byKind = k => pages.filter(p => p.kind === k && p.result);

  // ---- CUB 低圧幹線
  const cubPanels = [];
  for (const p of byKind('CUB')) {
    const incs = p.result.incoming || [];
    for (const c of p.result.circuits || []) {
      const group = nz(c.group) || nz(p.panel) || '低圧配電盤';
      const short = group.includes(' ') && !group.includes('試験') ? group.split(' ').pop() : group.replace(/\s/g, '');
      const inc = incs.find(i => nz(i.group) === nz(c.group)) || incs[0] || {};
      const [ph, v] = nominal(inc.system);
      if (!cubPanels.find(x => x['盤名'] === short)) {
        cubPanels.push({ '盤名': short, '種別': '低圧配電盤(CUB)', '電源(幹線番号)': '', '電気方式': `${ph} ${v}V`.trim(), '主幹': '', '幹線サイズ': '', '出典': src(p), '備考': [group, nz(inc.system)].filter(Boolean).join('  ') });
      }
      const name = nz(c.name);
      mid.trunk.push({
        '配電盤': short, '幹線番号': nz(c.no), '行先': name, '電気方式': ph, '電圧': v, '幹線サイズ': nz(c.cable), '許容電流(A)': '',
        'ブレーカ': nz(c.breaker), '容量': nz(c.capacity), '測定範囲': `${short} ～ ${name}`,
        '区分': (name === '予備' || c.kind === '予備') ? '予備' : (c.kind === 'space' ? '除外' : '幹線'), '備考': nz(c.remark), '出典': src(p),
      });
    }
    for (const x of p.result.checks || []) mid.checks.push({ '盤名': p.panel || '', 'ページ': src(p), '内容': x });
  }
  mid.trunk.forEach((r, i) => { r['並び'] = (i + 1) * 10; });

  // ---- 分電盤・動力盤
  const panelOrder = { '分電盤': [], '動力盤': [] };
  const perPanel = { '分電盤': new Map(), '動力盤': new Map() };
  for (const kind of ['動力盤', '分電盤']) {
    for (const p of byKind(kind)) {
      const panel = nz(p.panel) || nz(p.result.panel) || `${p.pdfName} p${p.page}`;
      if (!perPanel[kind].has(panel)) { perPanel[kind].set(panel, { items: [], inc: null, pages: [] }); panelOrder[kind].push(panel); }
      const P = perPanel[kind].get(panel);
      P.pages.push(p.page);
      for (const inc of p.result.incoming || []) if (inc && (inc.source || inc.cable || inc.system)) P.inc = mergeInc(P.inc, inc);
      for (const c of p.result.circuits || []) P.items.push({ ...c, _page: src(p), _seq: P.items.length });
      for (const x of p.result.checks || []) mid.checks.push({ '盤名': panel, 'ページ': src(p), '内容': x });
    }
  }

  for (const panel of panelOrder['動力盤']) {
    const P = perPanel['動力盤'].get(panel);
    const numbered = sortCircuits(P.items.filter(c => nz(c.no)));
    const items = numbered.concat(P.items.filter(c => !nz(c.no)));
    items.forEach((c, i) => {
      const name = nz(c.name);
      let cat = c.kind === 'space' ? '除外' : ((c.kind === '予備' || name === '予備') ? '予備' : '動力');
      if (name.includes('制御電源')) cat = '除外';
      mid.power.push({
        '並び': (i + 1) * 10, '盤名': panel, '回路番号': circNo(c), '負荷名称': name, '負荷記号': nz(c.symbol),
        '容量(kW)': nz(c.capacity), '電圧(V)': nz(c.volt), 'ケーブル': nz(c.cable), '許容電流(A)': '', 'ブレーカ': nz(c.breaker),
        '負荷電流(A)': nz(c.max_current), '区分': cat, '備考': [nz(c.remark)].filter(Boolean).join('; '), '出典': c._page,
      });
    });
  }

  for (const panel of panelOrder['分電盤']) {
    const P = perPanel['分電盤'].get(panel);
    const items = sortCircuits(P.items);
    const cats = items.map(c => classifyLight(c.name, c.kind));
    const votes = new Map();
    items.forEach((c, i) => { if (['電灯', 'コンセント'].includes(cats[i])) { const g = grpOf(c); votes.set(g, [...(votes.get(g) || []), cats[i]]); } });
    items.forEach((c, i) => {
      let cat = cats[i], rem = nz(c.remark);
      const name = nz(c.name), vs = votes.get(grpOf(c)) || [];
      if (cat === 'その他' && name && !EQUIP.test(name) && vs.length >= 3) {
        const top = ['電灯', 'コンセント'].sort((a, b) => vs.filter(x => x === b).length - vs.filter(x => x === a).length)[0];
        if (vs.filter(x => x === top).length / vs.length >= 0.8) { cat = top; rem = ['区分推定', rem].filter(Boolean).join('; '); }
      }
      if (cat === 'その他' && !name) rem = ['図面に負荷名称の記載なし', rem].filter(Boolean).join('; ');
      mid.light.push({
        '並び': (i + 1) * 10, '盤名': panel, '回路番号': circNo(c), '負荷名称': name, '容量(VA)': nz(c.capacity),
        '電圧(V)': nz(c.volt), 'ブレーカ': nz(c.breaker), '区分': cat, '備考': rem, '出典': c._page,
      });
    });
  }

  // 同じ盤内で番号が重複（別系統）
  for (const rows of [mid.light, mid.power]) {
    const seen = new Map();
    for (const r of rows) if (r['回路番号'] && !r['回路番号'].startsWith('(')) {
      const k = r['盤名'] + '\u0000' + r['回路番号'];
      seen.set(k, [...(seen.get(k) || []), r]);
    }
    for (const rs of seen.values()) if (rs.length > 1) for (const r of rs) r['備考'] = [`番号重複（別系統: ${r['出典']}）`, r['備考']].filter(Boolean).join('; ');
  }

  // ---- 盤一覧（CUB の幹線と照合）
  const feederOf = new Map();
  for (const t of mid.trunk) for (const d of t['行先'].split(/[,、]/)) if (!feederOf.has(nz(d))) feederOf.set(nz(d), t);
  mid.panels.push(...cubPanels);
  for (const kind of ['動力盤', '分電盤']) {
    for (const panel of panelOrder[kind]) {
      const P = perPanel[kind].get(panel);
      const inc = { ...(P.inc || {}) };
      const parts = k => nz(inc[k]).split(' / ').filter(Boolean);
      let dcnote = '';
      if (parts('source').length > 1 && parts('source').some(x => x.includes('AC-DC'))) {
        dcnote = '非常照明 DC100V（AC-DC-EC100V）併設';
        inc.source = parts('source').filter(x => !x.includes('AC-DC')).join(' / ');
        inc.cable = parts('cable').filter(x => !x.startsWith('FP')).join(' / ');
        inc.system = parts('system').filter(x => !x.includes('AC-DC') && !x.includes('DC100')).join(' / ');
      }
      const fb = feederOf.get(panel) || [...feederOf.entries()].find(([k]) => k.startsWith(panel + '('))?.[1];
      const srcNo = nz(inc.source);
      mid.panels.push({
        '盤名': panel, '種別': kind, '電源(幹線番号)': srcNo || (fb ? fb['幹線番号'] : ''), '電気方式': nz(inc.system), '主幹': nz(inc.main_breaker),
        '幹線サイズ': nz(inc.cable) || (fb ? fb['幹線サイズ'] : ''), '出典': `p${P.pages.join(',')}`,
        '備考': [fb ? `CUB: ${fb['配電盤']} ${fb['幹線番号']} ${fb['幹線サイズ']}` : '', dcnote].filter(Boolean).join('; '),
      });
      if (fb && srcNo) {
        const norm = t => nz(t).replace(/\s|E\d+(\.\d+)?sq/g, '').replace(/×/g, 'x').toUpperCase();
        const srcs = srcNo.split(' / ').filter(x => !x.includes('AC-DC'));
        const cabs = nz(inc.cable).split(' / ').map(norm);
        if (srcs.length && !srcs.join(' ').includes(fb['幹線番号'])) mid.checks.push({ '盤名': panel, 'ページ': '盤一覧', '内容': `電源番号が不一致: CUB図 ${fb['幹線番号']}（${fb['配電盤']}）／盤図 ${srcs[0]}` });
        else if (!cabs.includes(norm(fb['幹線サイズ']))) mid.checks.push({ '盤名': panel, 'ページ': '盤一覧', '内容': `幹線サイズが不一致: CUB図 ${fb['幹線番号']} ${fb['幹線サイズ']}／盤図 ${nz(inc.cable)}` });
      }
    }
  }
  return mid;
}

// ---------------------------------------------------------------- xlsx 読み書き
function cellText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.text !== undefined) return String(v.text);
    if (v.result !== undefined) return String(v.result);
    return '';
  }
  if (typeof v === 'number' && Number.isInteger(v)) return String(v);
  return String(v);
}

export async function readMidXlsx(buf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const mid = emptyMid();
  const common = {};
  const cws = wb.getWorksheet('共通');
  if (cws) cws.eachRow((row, r) => { if (r > 1) { const k = cellText(row.getCell(1).value).trim(); if (k) common[k] = cellText(row.getCell(2).value).trim(); } });
  for (const [key, name] of Object.entries(SHEETS)) {
    const ws = wb.getWorksheet(name);
    if (!ws) continue;
    const hdr = [];
    ws.getRow(1).eachCell({ includeEmpty: true }, (c, i) => { hdr[i] = cellText(c.value).trim(); });
    ws.eachRow((row, r) => {
      if (r === 1) return;
      const d = {}; let any = false;
      hdr.forEach((h, i) => { if (!h) return; const t = cellText(row.getCell(i).value).trim(); d[h] = t; if (t) any = true; });
      if (any) mid[key].push(d);
    });
  }
  return { mid, common };
}

export async function writeMidXlsx(mid, common, sources) {
  const wb = new ExcelJS.Workbook();
  const font = { name: 'Meiryo UI', size: 10 };
  const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
  const border = { left: thin, right: thin, top: thin, bottom: thin };
  const guide = [
    '中間リスト（添削用）の使い方', '',
    '1. 各シートの内容を盤図と見比べて直してください（黄色の「区分」列で出力先を切り替えます）。',
    '   ・低圧幹線 → B02 / B03（区分＝予備・除外は出力しない）',
    '   ・動力回路 → B04（除外以外）/ B05（予備・除外は出力しない）',
    '   ・電灯コンセント → B06（除外以外）/ B07（区分＝電灯）/ B08（区分＝コンセント）',
    '2. 行の追加・削除は自由です。並び順は「盤一覧」の順 → 各シートの「並び」列の昇順です。',
    '3. 許容電流(A) を記入すると B03・B05 の（　A）欄に入ります。',
    '4. 「共通」シートの立会者・実施者を入れると全ページのヘッダに入ります。',
    '5. Webアプリの「中間リスト」→「xlsxを読み込む」でこのファイルを戻せます。',
    '', '「要確認」シート: 図面の判読に自信がなかった箇所・図面どうしの食い違いの一覧です。',
    sources ? `元図面: ${sources}` : '',
  ];
  const g = wb.addWorksheet('使い方');
  guide.forEach((t, i) => { const c = g.getCell(i + 1, 1); c.value = t; c.font = { ...font, size: 11, bold: i === 0 }; });
  g.getColumn(1).width = 120;
  const cws = wb.addWorksheet('共通');
  [['項目', '値'], ['工事名称', common['工事名称'] || ''], ['立会者', common['立会者'] || ''], ['実施者', common['実施者'] || ''], ['年号', common['年号'] || '令和']]
    .forEach((r, i) => r.forEach((v, j) => { const c = cws.getCell(i + 1, j + 1); c.value = v || null; c.font = { ...font, bold: i === 0 }; }));
  cws.getColumn(1).width = 14; cws.getColumn(2).width = 60;
  const widths = {
    panels: [18, 14, 14, 22, 26, 18, 26, 44], trunk: [6, 18, 9, 24, 8, 9, 14, 9, 30, 10, 36, 8, 20, 14],
    power: [6, 10, 8, 26, 14, 9, 8, 16, 9, 30, 9, 8, 40, 16], light: [6, 14, 8, 34, 9, 8, 34, 10, 24, 16], checks: [14, 22, 120],
  };
  for (const key of ['panels', 'trunk', 'power', 'light', 'checks']) {
    const ws = wb.addWorksheet(SHEETS[key], { views: [{ state: 'frozen', xSplit: key === 'checks' || key === 'panels' ? 1 : 2, ySplit: 1 }] });
    const cols = COLS[key];
    cols.forEach((h, j) => {
      const c = ws.getCell(1, j + 1);
      c.value = h; c.font = { ...font, bold: true }; c.border = border;
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      ws.getColumn(j + 1).width = widths[key][j];
    });
    ws.getRow(1).height = 30;
    mid[key].forEach((r, i) => {
      cols.forEach((h, j) => {
        const c = ws.getCell(i + 2, j + 1);
        let v = r[h] ?? '';
        if (h === '並び' && v !== '' && Number.isFinite(+v)) v = +v;
        c.value = v === '' ? null : v; c.font = font; c.border = border;
        c.alignment = { vertical: 'middle', wrapText: ['負荷名称', 'ブレーカ', '備考', '測定範囲', '内容'].includes(h) };
        if (['予備', '除外'].includes(r['区分'])) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDEDED' } };
        else if (h === '区分') c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      });
    });
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(2, mid[key].length + 1), column: cols.length } };
    if (CATS[key]) {
      const j = cols.indexOf('区分') + 1;
      for (let i = 2; i <= mid[key].length + 200; i++) ws.getCell(i, j).dataValidation = { type: 'list', allowBlank: true, formulae: [`"${CATS[key].join(',')}"`] };
    }
  }
  const out = await wb.xlsx.writeBuffer();
  return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
