import { kvGet, kvSet, kvDel, keyStore } from './store.js';
import { openPdf, renderFit, pageText, extractionImages, classifyImages, pageSize } from './pdfpages.js';
import { MODELS, USD_JPY, costUSD, classifyBatch, extractPage, describeError, estimateUSD } from './ai.js';
import { buildMid, readMidXlsx, writeMidXlsx, emptyMid, COLS, SHEETS, CATS } from './mid.js';
import { FORM_DEFS, guessFormCode, makeForm } from './forms.js';
/* global JSZip */

const $ = id => document.getElementById(id);
const esc = t => String(t ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const KINDS = ['対象外', '分電盤', '動力盤', 'CUB'];
const yen = usd => `$${usd.toFixed(2)}（約${Math.round(usd * USD_JPY).toLocaleString()}円）`;

let state = {
  settings: { model: 'claude-opus-5', effort: 'high', clsModel: 'claude-opus-5', testMode: false },
  common: { '工事名称': '', '立会者': '', '実施者': '', '年号': '令和' },
  pdfs: [],      // {key, name, size, numPages}
  pages: {},     // pdfKey -> { [page]: {kind, panel, drawing, manual, cls, status, result, error, usd} }
  mid: emptyMid(),
  spent: 0,
};
const docs = new Map();         // pdfKey -> pdf.js document（PDF本体は保存しない）
const thumbURL = new Map();     // `${key}:${page}` -> objectURL
let tpls = {};                  // code -> {name}
let curPdf = null;
let midTab = 'light', midPanel = '', midSearch = '';
let running = null;

// ------------------------------------------------------------------ テスト用データ（APIキーなしで試す）
// 読み取り済みデータは、テスト用データのZIPから読み込んで端末に保存したもの（sample:results）か、
// このPC版に同梱した sample/ のもの。公開サイトには sample/ を置かない。
let sampleCache = null;
async function hasSample() {
  try { const r = await fetch('sample/templates.json', { cache: 'no-store' }); return r.ok; } catch { return false; }
}
async function sampleResults() {
  if (sampleCache) return sampleCache;
  const stored = await kvGet('sample:results');
  if (stored) return (sampleCache = stored);
  try { const r = await fetch('sample/results.json'); if (r.ok) return (sampleCache = await r.json()); } catch { /* なし */ }
  return null;
}
async function samplePagesFor(pdfName) {
  const s = await sampleResults();
  if (!s) return null;
  const n = pdfName.normalize('NFKC');
  const f = s.files.find(x => n.includes(x.match));
  return f ? f.pages : null;
}
const testMode = () => !!state.settings.testMode;
function markTest() {
  const h = document.querySelector('.brand h1');
  const b = h.querySelector('.testbadge'); if (b) b.remove();
  if (testMode()) h.insertAdjacentHTML('beforeend', '<span class="testbadge">テストモード</span>');
}
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ------------------------------------------------------------------ 保存
let saveTimer = null;
function save() { clearTimeout(saveTimer); saveTimer = setTimeout(() => kvSet('project', state).catch(e => console.error(e)), 500); }
async function load() {
  const p = await kvGet('project');
  if (p) state = { ...state, ...p, settings: { ...state.settings, ...(p.settings || {}) }, common: { ...state.common, ...(p.common || {}) }, mid: { ...emptyMid(), ...(p.mid || {}) } };
  for (const code of Object.keys(FORM_DEFS)) { const t = await kvGet('tpl:' + code); if (t) tpls[code] = { name: t.name }; }
  // 途中で閉じた読み取りは未読取に戻す
  for (const pg of Object.values(state.pages)) for (const ps of Object.values(pg)) if (ps.status === 'run') ps.status = ps.result ? 'done' : '';
}
function pstate(key, n) {
  const pg = state.pages[key] || (state.pages[key] = {});
  return pg[n] || (pg[n] = { kind: '対象外', panel: '', status: '' });
}

function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
}

// ------------------------------------------------------------------ 画面切替
function showStep(step) {
  document.querySelectorAll('.steps button').forEach(b => b.classList.toggle('on', b.dataset.step === step));
  document.querySelectorAll('.step').forEach(s => s.classList.toggle('on', s.id === 's-' + step));
  $('actionbar').classList.toggle('hidden', step !== 'pdf' || !!running);
  if (step === 'pdf') { renderPdfList(); renderThumbs(); updateActionbar(); }
  if (step === 'mid') renderMid();
  if (step === 'out') renderOut();
  window.scrollTo(0, 0);
}
$('steps').addEventListener('click', e => { const b = e.target.closest('button'); if (b) showStep(b.dataset.step); });

// ------------------------------------------------------------------ ① 設定
function initSettings() {
  for (const id of ['model', 'clsModel']) {
    $(id).innerHTML = Object.entries(MODELS).map(([k, m]) => `<option value="${k}">${esc(m.label)}</option>`).join('');
    $(id).value = state.settings[id];
    $(id).onchange = () => { state.settings[id] = $(id).value; save(); updateActionbar(); };
  }
  $('effort').value = state.settings.effort;
  $('effort').onchange = () => { state.settings.effort = $('effort').value; save(); };
  $('testMode').checked = testMode();
  $('testMode').onchange = () => { state.settings.testMode = $('testMode').checked; save(); updateActionbar(); markTest(); };
  $('loadSample').onclick = async () => {
    if (hasMid() && !confirm('いまの中間リストを、テスト用の中間リストで置き換えます。よろしいですか？')) return;
    const st = $('sampleStatus');
    st.textContent = '読み込み中…'; st.className = 'status';
    try {
      const names = await (await fetch('sample/templates.json')).json();
      for (const [code, name] of Object.entries(names)) {
        const buf = await (await fetch(`sample/${code}.xlsx`)).arrayBuffer();
        await kvSet('tpl:' + code, { name, buf });
        tpls[code] = { name };
      }
      const { mid, common } = await readMidXlsx(await (await fetch('sample/mid.xlsx')).arrayBuffer());
      state.mid = mid; midPanel = '';
      for (const k of Object.keys(state.common)) if (common[k]) { state.common[k] = common[k]; if ($('c-' + k)) $('c-' + k).value = common[k]; }
      save(); renderTpl(); $('testModeRow').classList.remove('hidden');
      st.textContent = `読み込みました（社内書式 ${Object.keys(names).length}種、中間リスト: 盤${mid.panels.length}・低圧幹線${mid.trunk.length}・動力${mid.power.length}・電灯コンセント${mid.light.length}行）。③ 中間リストへ進んでください。`;
      st.className = 'status ok';
    } catch (e) { console.error(e); st.textContent = '読み込めませんでした: ' + (e.message || e); st.className = 'status err'; }
  };
  // テスト用データのZIP（社内書式・中間リスト・読み取り済みデータ・盤図PDF）をまとめて読み込む
  $('sampleZip').onchange = async e => {
    const file = e.target.files[0]; e.target.value = '';
    if (!file) return;
    if (hasMid() && !confirm('いまの中間リストを、ZIPの中間リストで置き換えます。よろしいですか？')) return;
    const st = $('sampleStatus');
    st.textContent = '読み込み中…（大きなZIPは少し時間がかかります）'; st.className = 'status';
    try {
      const zip = await JSZip.loadAsync(file);
      const got = { tpl: 0, mid: null, res: false, pdf: [] };
      for (const [path, entry] of Object.entries(zip.files)) {
        if (entry.dir) continue;
        const base = path.split(/[\\/]/).pop();   // Windows で作ったZIPは区切りが \ のことがある
        if (/\.xlsx$/i.test(base)) {
          const code = guessFormCode(base);
          if (code) { await kvSet('tpl:' + code, { name: base, buf: await entry.async('arraybuffer') }); tpls[code] = { name: base }; got.tpl++; }
          else if (base.includes('中間リスト')) {
            const { mid, common } = await readMidXlsx(await entry.async('arraybuffer'));
            state.mid = mid; midPanel = ''; got.mid = mid;
            for (const k of Object.keys(state.common)) if (common[k]) { state.common[k] = common[k]; if ($('c-' + k)) $('c-' + k).value = common[k]; }
          }
        } else if (/\.json$/i.test(base)) {
          const j = JSON.parse(await entry.async('string'));
          if (j && Array.isArray(j.files)) { await kvSet('sample:results', j); sampleCache = j; got.res = true; }
        } else if (/\.pdf$/i.test(base)) {
          got.pdf.push(new File([await entry.async('blob')], base, { type: 'application/pdf' }));
        }
      }
      if (got.pdf.length) await addPdfFiles(got.pdf);
      save(); renderTpl();
      if (got.res) $('testModeRow').classList.remove('hidden');
      const m = got.mid;
      st.textContent = `読み込みました: 社内書式 ${got.tpl}種` + (m ? `、中間リスト（盤${m.panels.length}・低圧幹線${m.trunk.length}・動力${m.power.length}・電灯コンセント${m.light.length}行）` : '') +
        (got.pdf.length ? `、盤図PDF ${got.pdf.length}件` : '') + (got.res ? '、読み取り済みデータ（テストモードが使えます）' : '') + '。';
      st.className = 'status ok';
    } catch (err) { console.error(err); st.textContent = '読み込めませんでした: ' + (err.message || err); st.className = 'status err'; }
  };
  $('apiKey').value = keyStore.get();
  $('rememberKey').checked = keyStore.remembered();
  const keyMsg = () => { const k = keyStore.get(); $('keyStatus').textContent = k ? `設定済み（…${k.slice(-4)}）` : '未設定'; $('keyStatus').className = 'status ' + (k ? 'ok' : 'err'); };
  $('saveKey').onclick = () => { keyStore.set($('apiKey').value.trim(), $('rememberKey').checked); keyMsg(); };
  $('rememberKey').onchange = () => { if (keyStore.get()) keyStore.set(keyStore.get(), $('rememberKey').checked); };
  keyMsg();
  for (const k of Object.keys(state.common)) {
    const el = $('c-' + k);
    if (!el) continue;
    el.value = state.common[k] || '';
    el.oninput = () => { state.common[k] = el.value; save(); };
  }
  $('tplInput').onchange = async e => {
    const msgs = [];
    for (const f of e.target.files) {
      const code = guessFormCode(f.name);
      if (!code) { msgs.push(`${f.name}: 対象外（B02〜B08 ではない）`); continue; }
      await kvSet('tpl:' + code, { name: f.name, buf: await f.arrayBuffer() });
      tpls[code] = { name: f.name };
    }
    e.target.value = '';
    renderTpl(msgs);
  };
  $('resetProject').onclick = async () => {
    if (!confirm('ページ判定・読み取り結果・中間リストを消去して、新しい現場を始めます。よろしいですか？（書式とAPIキーは残ります）')) return;
    state.pdfs = []; state.pages = {}; state.mid = emptyMid(); state.spent = 0; docs.clear(); curPdf = null;
    await kvSet('project', state); renderPdfList(); renderThumbs(); alert('消去しました。');
  };
  renderTpl();
}
function renderTpl(msgs = []) {
  $('tplList').innerHTML = Object.entries(FORM_DEFS).map(([k, d]) =>
    `<li class="${tpls[k] ? '' : 'miss'}"><b>${k}</b> ${tpls[k] ? esc(tpls[k].name) : '未設定'}</li>`).join('') +
    msgs.map(m => `<li class="miss">${esc(m)}</li>`).join('');
}

// ------------------------------------------------------------------ ② 盤図PDF
async function addPdfFiles(files) {
  for (const f of files) {
    const key = `${f.name}|${f.size}`;
    try {
      const doc = await openPdf(f);
      docs.set(key, doc);
      if (!state.pdfs.find(p => p.key === key)) state.pdfs.push({ key, name: f.name, size: f.size, numPages: doc.numPages });
      state.pages[key] = state.pages[key] || {};
      curPdf = key;
    } catch (err) { alert(`${f.name} を開けませんでした: ${err.message || err}`); }
  }
  save(); renderPdfList(); renderThumbs(); updateActionbar();
}
$('pdfInput').onchange = async e => {
  const files = [...e.target.files];
  e.target.value = '';
  await addPdfFiles(files);
};

function pdfCounts(key) {
  const pg = state.pages[key] || {};
  const c = { target: 0, done: 0, err: 0, cls: 0 };
  for (const ps of Object.values(pg)) {
    if (ps.cls || ps.manual) c.cls++;
    if (ps.kind !== '対象外') { c.target++; if (ps.status === 'done') c.done++; if (ps.status === 'err') c.err++; }
  }
  return c;
}
function renderPdfList() {
  $('pdfList').innerHTML = state.pdfs.map(p => {
    const c = pdfCounts(p.key);
    return `<div class="pdfitem"><b>${esc(p.name)}</b><span class="chip">${p.numPages}ページ</span>
      <span class="chip">判定済 ${c.cls}</span><span class="chip">読み取り対象 ${c.target}</span><span class="chip">読取済 ${c.done}</span>
      ${c.err ? `<span class="chip" style="color:var(--warn)">エラー ${c.err}</span>` : ''}
      ${docs.has(p.key) ? '' : '<span class="miss">PDF未読込（同じPDFを選び直してください）</span>'}
      <button class="btn small" data-show="${esc(p.key)}">表示</button><button class="btn small danger" data-del="${esc(p.key)}">外す</button></div>`;
  }).join('') || '<p class="note">まだPDFがありません。</p>';
  $('pdfPanel').classList.toggle('hidden', !state.pdfs.length);
  $('pdfSel').innerHTML = state.pdfs.map(p => `<option value="${esc(p.key)}">${esc(p.name)}</option>`).join('');
  if (!curPdf || !state.pdfs.find(p => p.key === curPdf)) curPdf = state.pdfs[0] ? state.pdfs[0].key : null;
  if (curPdf) $('pdfSel').value = curPdf;
}
$('pdfList').onclick = e => {
  const s = e.target.closest('[data-show]'), d = e.target.closest('[data-del]');
  if (s) { curPdf = s.dataset.show; renderPdfList(); renderThumbs(); }
  if (d && confirm('このPDFとページ判定・読み取り結果を外しますか？')) {
    state.pdfs = state.pdfs.filter(p => p.key !== d.dataset.del); delete state.pages[d.dataset.del]; docs.delete(d.dataset.del);
    save(); renderPdfList(); renderThumbs(); updateActionbar();
  }
};
$('pdfSel').onchange = () => { curPdf = $('pdfSel').value; renderThumbs(); };
$('pageFilter').onchange = () => renderThumbs();

function badge(ps) {
  if (ps.status === 'run') return '<span class="badge run">読取中…</span>';
  if (ps.status === 'done') return `<span class="badge done">読取済 ${(ps.result && ps.result.circuits || []).length}回路</span>`;
  if (ps.status === 'err') return `<span class="badge err" title="${esc(ps.error)}">エラー</span>`;
  if (ps.kind !== '対象外') return '<span class="badge">未読取</span>';
  return (ps.cls || ps.manual) ? '<span class="badge">対象外</span>' : '<span class="badge">未判定</span>';
}
function cardHTML(key, n) {
  const ps = pstate(key, n);
  const again = (ps.status === 'done' || ps.status === 'err') ? `<button class="btn small" data-reread="${n}">再読取</button>` : '';
  return `<div class="th k-${ps.kind}" data-p="${n}">
    <div class="img" data-view="${n}">${thumbURL.has(key + ':' + n) ? `<img src="${thumbURL.get(key + ':' + n)}" alt="">` : `<span>p${n}</span>`}</div>
    <div class="meta"><b>p${n}</b>${badge(ps)}</div>
    <select data-kind="${n}">${KINDS.map(k => `<option${k === ps.kind ? ' selected' : ''}>${k}</option>`).join('')}</select>
    <input data-panel="${n}" type="text" placeholder="盤名" value="${esc(ps.panel)}">${again}</div>`;
}
let thumbObserver = null;
function renderThumbs() {
  const box = $('thumbs');
  if (thumbObserver) thumbObserver.disconnect();
  const p = state.pdfs.find(x => x.key === curPdf);
  if (!p) { box.innerHTML = ''; return; }
  const f = $('pageFilter').value;
  const list = [];
  for (let n = 1; n <= p.numPages; n++) {
    const ps = pstate(p.key, n);
    if (f === 'target' && ps.kind === '対象外') continue;
    if (f === 'todo' && (ps.cls || ps.manual)) continue;
    list.push(n);
  }
  box.innerHTML = list.map(n => cardHTML(p.key, n)).join('') || '<p class="note">該当するページはありません。</p>';
  if (!docs.has(p.key)) return;
  thumbObserver = new IntersectionObserver(ents => {
    for (const en of ents) if (en.isIntersecting) { thumbObserver.unobserve(en.target); queueThumb(p.key, +en.target.dataset.view, en.target); }
  }, { rootMargin: '400px' });
  box.querySelectorAll('.img').forEach(el => { if (!el.querySelector('img')) thumbObserver.observe(el); });
}
const thumbQ = []; let thumbBusy = 0;
function queueThumb(key, n, el) { thumbQ.push({ key, n, el }); pumpThumbs(); }
async function pumpThumbs() {
  while (thumbBusy < 2 && thumbQ.length) {
    const { key, n, el } = thumbQ.shift();
    thumbBusy++;
    (async () => {
      try {
        const id = key + ':' + n;
        if (!thumbURL.has(id)) {
          const cv = await renderFit(docs.get(key), n, 360);
          const blob = await new Promise(r => cv.toBlob(r, 'image/jpeg', 0.7));
          cv.width = cv.height = 0;
          thumbURL.set(id, URL.createObjectURL(blob));
        }
        if (el.isConnected) el.innerHTML = `<img src="${thumbURL.get(id)}" alt="">`;
      } catch (e) { console.warn(e); }
      thumbBusy--; pumpThumbs();
    })();
  }
}
function refreshCard(key, n) {
  if (key !== curPdf) return;
  const el = $('thumbs').querySelector(`.th[data-p="${n}"]`);
  if (el) { el.outerHTML = cardHTML(key, n); }
}
$('thumbs').addEventListener('change', e => {
  const t = e.target;
  if (t.dataset.kind) { const ps = pstate(curPdf, +t.dataset.kind); ps.kind = t.value; ps.manual = true; t.closest('.th').className = `th k-${ps.kind}`; save(); updateActionbar(); }
});
$('thumbs').addEventListener('input', e => {
  const t = e.target;
  if (t.dataset.panel) { const ps = pstate(curPdf, +t.dataset.panel); ps.panel = t.value.trim(); ps.manual = true; save(); }
});
$('thumbs').addEventListener('click', async e => {
  const v = e.target.closest('[data-view]'), r = e.target.closest('[data-reread]');
  if (r) { const ps = pstate(curPdf, +r.dataset.reread); ps.status = ''; ps.result = null; refreshCard(curPdf, +r.dataset.reread); save(); updateActionbar(); return; }
  if (v && docs.has(curPdf)) openViewer(curPdf, +v.dataset.view);
});
$('rgApply').onclick = () => {
  const a = +$('rgFrom').value, b = +$('rgTo').value || a, p = state.pdfs.find(x => x.key === curPdf);
  if (!p || !a) return;
  for (let n = Math.max(1, a); n <= Math.min(p.numPages, b); n++) {
    const ps = pstate(curPdf, n); ps.kind = $('rgKind').value; ps.manual = true;
    if ($('rgPanel').value.trim()) ps.panel = $('rgPanel').value.trim();
  }
  save(); renderThumbs(); updateActionbar(); renderPdfList();
};

async function openViewer(key, n) {
  const d = $('viewer');
  $('vTitle').textContent = `${state.pdfs.find(p => p.key === key).name}  p${n}`;
  $('vImg').src = '';
  d.showModal();
  const cv = await renderFit(docs.get(key), n, 2400);
  const blob = await new Promise(r => cv.toBlob(r, 'image/jpeg', 0.85));
  cv.width = cv.height = 0;
  const old = $('vImg').src; $('vImg').src = URL.createObjectURL(blob); if (old.startsWith('blob:')) URL.revokeObjectURL(old);
}
$('vClose').onclick = () => $('viewer').close();
$('vZoom').onclick = () => $('viewer').querySelector('.vbody').classList.toggle('zoom');

// ---- 実行バー
let wakeLock = null;
async function keepAwake(on) {
  // 読み取り中に iPad の画面が消えると Safari が通信を止めるため、画面を点けたままにする
  try {
    if (on && 'wakeLock' in navigator) wakeLock = await navigator.wakeLock.request('screen');
    else if (!on && wakeLock) { await wakeLock.release(); wakeLock = null; }
  } catch { /* 非対応の端末では何もしない */ }
}
document.addEventListener('visibilitychange', () => { if (running && document.visibilityState === 'visible') keepAwake(true); });
function startRun(title) {
  keepAwake(true);
  running = { ctrl: new AbortController(), title };
  $('runTitle').textContent = title; $('runText').textContent = ''; $('runBar').style.width = '0%';
  $('runbar').classList.remove('hidden'); $('actionbar').classList.add('hidden');
  return running.ctrl.signal;
}
function runProgress(done, total, text) { $('runBar').style.width = `${Math.round(done / Math.max(1, total) * 100)}%`; $('runText').textContent = text; }
function endRun() { keepAwake(false); running = null; $('runbar').classList.add('hidden'); updateActionbar(); }
$('runStop').onclick = () => { if (running) running.ctrl.abort(); };
window.addEventListener('beforeunload', e => { if (running) { e.preventDefault(); e.returnValue = ''; } });

function needKey() {
  const k = keyStore.get();
  if (!k) { alert('① 設定 で Claude のAPIキーを入れてください。'); showStep('settings'); }
  return k;
}

async function pool(items, n, fn, signal) {
  let i = 0;
  const worker = async () => { while (i < items.length && !signal.aborted) { const it = items[i++]; await fn(it); } };
  await Promise.all(Array.from({ length: Math.min(n, items.length) }, worker));
}

// ---- AIでページ判定
async function classifyTest(key, p) {
  const sp = await samplePagesFor(p.name);
  if (!sp) { alert('テストモード: このPDFの読み取り済みデータがありません。① 設定でテスト用データ（ZIP）を読み込み、ZIPに入っている盤図PDFで試してください。'); return; }
  const signal = startRun('ページ判定（テストモード・0円）');
  let n = 0;
  for (let pg = 1; pg <= p.numPages && !signal.aborted; pg++) {
    const ps = pstate(key, pg);
    if (!ps.manual) {
      const s = sp[pg];
      ps.kind = s ? s.kind : '対象外'; ps.panel = s ? s.panel : ps.panel; ps.cls = true;
      if (s) n++;
      refreshCard(key, pg);
    }
    if (pg % 20 === 0) { runProgress(pg, p.numPages, `${pg}/${p.numPages}ページ`); await sleep(30); }
  }
  endRun(); save(); renderPdfList(); renderThumbs();
  alert(`テストモード: ${p.numPages}ページ中 ${n}ページを読み取り対象にしました（0円）。`);
}

$('classifyBtn').onclick = async () => {
  const key = curPdf, p = state.pdfs.find(x => x.key === key);
  if (!p || !docs.has(key)) { alert('PDFを読み込んでください。'); return; }
  if (testMode()) return classifyTest(key, p);
  const apiKey = needKey(); if (!apiKey) return;
  const todo = [];
  for (let n = 1; n <= p.numPages; n++) { const ps = pstate(key, n); if (!ps.cls && !ps.manual) todo.push(n); }
  if (!todo.length) { if (!confirm('全ページ判定済みです。もう一度判定しますか？（手で直した設定は残します）')) return; for (let n = 1; n <= p.numPages; n++) if (!pstate(key, n).manual) todo.push(n); }
  const m = state.settings.clsModel, price = MODELS[m];
  const est = todo.length * (1400 * price.in + 150 * price.out) / 1e6;
  if (!confirm(`${todo.length}ページをAIで判定します。\n見積り: ${yen(est)}\nモデル: ${price.label}`)) return;
  const signal = startRun('ページ判定');
  const batches = []; for (let i = 0; i < todo.length; i += 6) batches.push(todo.slice(i, i + 6));
  let done = 0, spent = 0, errs = 0;
  await pool(batches, 2, async b => {
    try {
      const items = [];
      for (const n of b) items.push({ page: n, images: await classifyImages(docs.get(key), n), text: await pageText(docs.get(key), n) });
      const res = await classifyBatch({ apiKey, model: m, items, signal });
      for (const r of res.json.pages || []) {
        if (!b.includes(r.page)) continue;
        const ps = pstate(key, r.page);
        if (ps.manual) continue;
        ps.kind = r.kind; ps.drawing = r.drawing; ps.cls = true;
        if (r.kind !== '対象外' || !ps.panel) ps.panel = r.panel || ps.panel;
        refreshCard(key, r.page);
      }
      const usd = costUSD(m, res.usage); spent += usd; state.spent += usd;
    } catch (e) {
      errs++;
      if (signal.aborted) return;
      console.error(e);
      const msg = describeError(e);
      if (/401|403/.test(msg)) { alert(msg); running && running.ctrl.abort(); }
    }
    done += b.length;
    runProgress(done, todo.length, `${done}/${todo.length}ページ　${yen(spent)}${errs ? `　失敗 ${errs}回` : ''}`);
    save();
  }, signal);
  endRun(); renderPdfList(); renderThumbs();
};

// ---- 読み取り
function targets(includeDone = false) {
  const out = [];
  for (const p of state.pdfs) {
    const pg = state.pages[p.key] || {};
    for (const [n, ps] of Object.entries(pg)) if (ps.kind !== '対象外' && (includeDone || ps.status !== 'done')) out.push({ key: p.key, n: +n, ps, pdf: p });
  }
  return out.sort((a, b) => state.pdfs.indexOf(a.pdf) - state.pdfs.indexOf(b.pdf) || a.n - b.n);
}
function updateActionbar() {
  const all = targets(true), todo = targets(false), loaded = todo.filter(t => docs.has(t.key));
  const est = estimateUSD(state.settings.model, loaded.length * 14e6, loaded.length);
  $('actText').innerHTML = `読み取り対象 <b>${all.length}</b>ページ（未読取 <b>${todo.length}</b>）` +
    (testMode() ? '<span class="testbadge">テストモード・0円</span>' : (loaded.length ? `　見積り ${yen(est)}` : '')) +
    (state.spent ? `　これまでの利用 ${yen(state.spent)}` : '');
  $('readBtn').disabled = !loaded.length || !!running;
  const onPdf = document.querySelector('#s-pdf').classList.contains('on');
  $('actionbar').classList.toggle('hidden', !onPdf || !!running || !state.pdfs.length);
}
async function readTest() {
  const list = targets(false).filter(t => docs.has(t.key));
  if (!confirm(`テストモード: ${list.length}ページを、同梱の読み取り済みデータで読み取ります（0円）。`)) return;
  const signal = startRun('読み取り（テストモード・0円）');
  let done = 0, miss = 0;
  for (const t of list) {
    if (signal.aborted) break;
    const sp = await samplePagesFor(t.pdf.name);
    const s = sp && sp[t.n];
    if (s) { t.ps.result = structuredClone(s.result); t.ps.status = 'done'; t.ps.error = ''; t.ps.usd = 0; if (!t.ps.panel) t.ps.panel = s.panel; }
    else { t.ps.status = 'err'; t.ps.error = 'テストモード: このページの読み取り済みデータがありません'; miss++; }
    refreshCard(t.key, t.n);
    done++;
    runProgress(done, list.length, `${done}/${list.length}ページ${miss ? `　データなし ${miss}` : ''}`);
    await sleep(40);
  }
  const aborted = signal.aborted;
  endRun(); renderPdfList(); save();
  if (!aborted && confirm(`読み取りが終わりました（データなし ${miss}ページ）。\n中間リストを作成しますか？${hasMid() ? '\n※いまの中間リストは上書きされます。' : ''}`)) { doBuildMid(); showStep('mid'); }
}

$('readBtn').onclick = async () => {
  if (testMode()) return readTest();
  const apiKey = needKey(); if (!apiKey) return;
  const list = targets(false).filter(t => docs.has(t.key));
  const miss = targets(false).length - list.length;
  const m = state.settings.model;
  const est = estimateUSD(m, list.length * 14e6, list.length);
  if (!confirm(`${list.length}ページを読み取ります${miss ? `（PDF未読込のため ${miss}ページは除外）` : ''}。\n見積り: ${yen(est)}\nモデル: ${MODELS[m].label}／精度: ${$('effort').selectedOptions[0].text}\n\n大きな盤は1ページ数分かかります。読み取り中はこの画面を開いたままにしてください（ほかのアプリに切り替えると止まることがあります）。途中で中止しても、読み取り済みのページは残ります。`)) return;
  const signal = startRun('図面の読み取り');
  let done = 0, spent = 0, errs = 0;
  runProgress(0, list.length, '準備中…');
  await pool(list, 2, async t => {
    const ps = t.ps;
    ps.status = 'run'; refreshCard(t.key, t.n);
    try {
      const images = await extractionImages(docs.get(t.key), t.n);
      const res = await extractPage({ apiKey, model: m, effort: state.settings.effort, kind: ps.kind, panelHint: ps.panel, pageLabel: `${t.pdf.name} p${t.n}`, images, signal });
      ps.result = res.json; ps.status = 'done'; ps.error = '';
      if (!ps.panel && res.json.panel) ps.panel = res.json.panel;
      const usd = costUSD(m, res.usage); ps.usd = usd; spent += usd; state.spent += usd;
    } catch (e) {
      ps.status = signal.aborted ? (ps.result ? 'done' : '') : 'err';
      ps.error = describeError(e);
      if (!signal.aborted) { errs++; console.error(e); if (/401|403/.test(ps.error)) { alert(ps.error); running && running.ctrl.abort(); } }
    }
    refreshCard(t.key, t.n);
    done++;
    runProgress(done, list.length, `${done}/${list.length}ページ　${yen(spent)}${errs ? `　エラー ${errs}` : ''}`);
    save();
  }, signal);
  const aborted = signal.aborted;
  endRun(); renderPdfList();
  const doneAll = targets(true).filter(t => t.ps.status === 'done').length;
  if (!aborted && doneAll && confirm(`読み取りが終わりました（エラー ${errs}ページ）。\n中間リストを作成しますか？${hasMid() ? '\n※いまの中間リストは上書きされます。' : ''}`)) { doBuildMid(); showStep('mid'); }
};

// ------------------------------------------------------------------ ③ 中間リスト
const hasMid = () => ['trunk', 'power', 'light'].some(k => state.mid[k].length);
function doBuildMid() {
  const list = [];
  for (const p of state.pdfs) {
    const pg = state.pages[p.key] || {};
    Object.keys(pg).map(Number).sort((a, b) => a - b).forEach(n => {
      const ps = pg[n];
      if (ps.kind !== '対象外' && ps.result) list.push({ pdfName: p.name, page: n, kind: ps.kind, panel: ps.panel || ps.result.panel, result: ps.result });
    });
  }
  state.mid = buildMid(list);
  midPanel = ''; save(); renderMid();
}
$('buildMid').onclick = () => {
  if (hasMid() && !confirm('読み取り結果から作り直します。中間リストで直した内容は消えます。よろしいですか？')) return;
  doBuildMid();
};
$('midInput').onchange = async e => {
  const f = e.target.files[0]; e.target.value = '';
  if (!f) return;
  if (hasMid() && !confirm('いまの中間リストを、選んだファイルの内容で置き換えます。よろしいですか？')) return;
  try {
    const { mid, common } = await readMidXlsx(await f.arrayBuffer());
    state.mid = mid;
    for (const k of Object.keys(state.common)) if (common[k]) { state.common[k] = common[k]; if ($('c-' + k)) $('c-' + k).value = common[k]; }
    midPanel = ''; save(); renderMid();
  } catch (err) { alert('読み込めませんでした: ' + (err.message || err)); }
};
$('midSave').onclick = async () => {
  const blob = await writeMidXlsx(state.mid, state.common, state.pdfs.map(p => p.name).join(' / '));
  download(blob, '中間リスト（添削用）.xlsx');
};

const PANEL_COL = { trunk: '配電盤', power: '盤名', light: '盤名', checks: '盤名', panels: '' };
function midRows() {
  const rows = state.mid[midTab] || [];
  const pc = PANEL_COL[midTab];
  const q = midSearch.toLowerCase();
  const idx = [];
  rows.forEach((r, i) => {
    if (pc && midPanel && r[pc] !== midPanel) return;
    if (q && !COLS[midTab].some(h => String(r[h] ?? '').toLowerCase().includes(q))) return;
    idx.push(i);
  });
  if (midTab !== 'panels' && midTab !== 'checks') idx.sort((a, b) => {
    const ra = rows[a], rb = rows[b];
    const po = p => { const i = state.mid.panels.findIndex(x => x['盤名'] === p); return i < 0 ? 1e6 : i; };
    return (po(ra[pc]) - po(rb[pc])) || String(ra[pc]).localeCompare(String(rb[pc])) || ((+ra['並び'] || 0) - (+rb['並び'] || 0)) || a - b;
  });
  return idx;
}
function renderMid() {
  const counts = k => state.mid[k].length;
  $('midTabs').innerHTML = Object.entries(SHEETS).map(([k, n]) => `<button data-tab="${k}" class="${k === midTab ? 'on' : ''}">${n} ${counts(k)}</button>`).join('');
  const pc = PANEL_COL[midTab];
  if (pc) {
    const names = [...new Set(state.mid[midTab].map(r => r[pc]).filter(Boolean))];
    if (midPanel && !names.includes(midPanel)) midPanel = '';
    if (!midPanel && names.length && state.mid[midTab].length > 300) midPanel = names[0];
    $('midPanel').innerHTML = `<option value="">すべての盤</option>` + names.map(n => `<option${n === midPanel ? ' selected' : ''}>${esc(n)}</option>`).join('');
    $('midPanel').classList.remove('hidden');
  } else $('midPanel').classList.add('hidden');
  $('midAdd').classList.toggle('hidden', midTab === 'checks');
  const cols = COLS[midTab];
  const rows = state.mid[midTab];
  const idx = midRows();
  const LIMIT = 500;
  const head = `<tr>${cols.map(h => `<th>${esc(h)}</th>`).join('')}<th></th></tr>`;
  const body = idx.slice(0, LIMIT).map(i => rowHTML(rows[i], i, cols)).join('');
  $('midTable').innerHTML = `<thead>${head}</thead><tbody>${body}</tbody>`;
  $('midInfo').textContent = !rows.length ? 'まだデータがありません。「読み取り結果から作成」か「xlsxを読み込む」を使ってください。'
    : `${idx.length}行を表示${idx.length > LIMIT ? `（先頭${LIMIT}行のみ。盤で絞り込んでください）` : ''}　／　全${rows.length}行`;
}
const WIDE = { '負荷名称': 16, '行先': 12, 'ブレーカ': 18, '備考': 16, '測定範囲': 18, '内容': 40, '電気方式': 12, '主幹': 16, '出典': 12, '電源(幹線番号)': 12, 'ケーブル': 10, '幹線サイズ': 10 };
function rowHTML(r, i, cols) {
  const off = ['予備', '除外'].includes(r['区分']);
  const cells = cols.map(h => {
    const v = r[h] ?? '';
    if (h === '区分' && CATS[midTab]) return `<td class="cat"><select data-i="${i}" data-h="${h}">${CATS[midTab].map(c => `<option${c === v ? ' selected' : ''}>${c}</option>`).join('')}${CATS[midTab].includes(v) ? '' : `<option selected>${esc(v)}</option>`}</select></td>`;
    const w = WIDE[h] || (h === '並び' ? 4 : 7);
    return `<td><input data-i="${i}" data-h="${esc(h)}" value="${esc(v)}" style="min-width:${w}em"></td>`;
  }).join('');
  const ops = midTab === 'checks' ? `<button data-op="del" data-i="${i}">×</button>` :
    `<button data-op="add" data-i="${i}" title="下に行を追加">＋</button><button data-op="up" data-i="${i}">↑</button><button data-op="down" data-i="${i}">↓</button><button data-op="del" data-i="${i}">×</button>`;
  return `<tr class="${off ? 'off' : ''}" data-i="${i}">${cells}<td class="ops">${ops}</td></tr>`;
}
$('midTabs').onclick = e => { const b = e.target.closest('button'); if (b) { midTab = b.dataset.tab; midPanel = ''; renderMid(); } };
$('midPanel').onchange = () => { midPanel = $('midPanel').value; renderMid(); };
$('midSearch').oninput = () => { midSearch = $('midSearch').value.trim(); renderMid(); };
$('midTable').addEventListener('input', e => {
  const t = e.target; if (t.dataset.i === undefined) return;
  state.mid[midTab][+t.dataset.i][t.dataset.h] = t.value; save();
});
$('midTable').addEventListener('change', e => {
  const t = e.target; if (t.dataset.h !== '区分') return;
  const r = state.mid[midTab][+t.dataset.i]; r['区分'] = t.value; save();
  t.closest('tr').classList.toggle('off', ['予備', '除外'].includes(t.value));
});
function renumber(panel) {
  const pc = PANEL_COL[midTab];
  const rows = state.mid[midTab].filter(r => r[pc] === panel).sort((a, b) => (+a['並び'] || 0) - (+b['並び'] || 0));
  rows.forEach((r, k) => { r['並び'] = (k + 1) * 10; });
}
$('midTable').addEventListener('click', e => {
  const b = e.target.closest('button[data-op]'); if (!b) return;
  const rows = state.mid[midTab], i = +b.dataset.i, r = rows[i], pc = PANEL_COL[midTab];
  if (b.dataset.op === 'del') { if (!confirm('この行を削除しますか？')) return; rows.splice(i, 1); }
  else if (midTab === 'panels') {
    if (b.dataset.op === 'add') rows.splice(i + 1, 0, Object.fromEntries(COLS.panels.map(h => [h, ''])));
    if (b.dataset.op === 'up' && i > 0) [rows[i - 1], rows[i]] = [rows[i], rows[i - 1]];
    if (b.dataset.op === 'down' && i < rows.length - 1) [rows[i + 1], rows[i]] = [rows[i], rows[i + 1]];
  } else {
    renumber(r[pc]);
    const same = rows.filter(x => x[pc] === r[pc]).sort((a, c) => a['並び'] - c['並び']);
    const k = same.indexOf(r);
    if (b.dataset.op === 'add') {
      const nr = Object.fromEntries(COLS[midTab].map(h => [h, '']));
      nr[pc] = r[pc]; nr['並び'] = r['並び'] + 5; nr['区分'] = CATS[midTab][0];
      if (midTab === 'trunk') { nr['電気方式'] = r['電気方式']; nr['電圧'] = r['電圧']; }
      if (midTab !== 'trunk') nr['電圧(V)'] = r['電圧(V)'];
      rows.splice(i + 1, 0, nr);
    }
    if (b.dataset.op === 'up' && k > 0) [r['並び'], same[k - 1]['並び']] = [same[k - 1]['並び'], r['並び']];
    if (b.dataset.op === 'down' && k < same.length - 1) [r['並び'], same[k + 1]['並び']] = [same[k + 1]['並び'], r['並び']];
    renumber(r[pc]);
  }
  save(); renderMid();
});
$('midAdd').onclick = () => {
  const rows = state.mid[midTab], pc = PANEL_COL[midTab];
  const nr = Object.fromEntries(COLS[midTab].map(h => [h, '']));
  if (pc) {
    nr[pc] = midPanel || (rows[rows.length - 1] || {})[pc] || '';
    const same = rows.filter(x => x[pc] === nr[pc]);
    nr['並び'] = same.reduce((m, x) => Math.max(m, +x['並び'] || 0), 0) + 10;
    if (CATS[midTab]) nr['区分'] = CATS[midTab][0];
  }
  rows.push(nr); save(); renderMid();
  const w = $('midTable').closest('.tablewrap'); w.scrollTop = w.scrollHeight;
};

// ------------------------------------------------------------------ ④ 書式出力
const FORM_SRC = { B02: 'trunk', B03: 'trunk', B04: 'power', B05: 'power', B06: 'light', B07: 'light', B08: 'light' };
function formRows(code) {
  const rows = state.mid[FORM_SRC[code]];
  const pc = code <= 'B03' ? '配電盤' : '盤名';
  const f = {
    B02: r => !['予備', '除外'].includes(r['区分']), B03: r => !['予備', '除外'].includes(r['区分']),
    B04: r => r['区分'] !== '除外', B05: r => !['予備', '除外'].includes(r['区分']),
    B06: r => r['区分'] !== '除外', B07: r => r['区分'] === '電灯', B08: r => r['区分'] === 'コンセント',
  }[code];
  const sel = rows.filter(f);
  return { n: sel.length, panels: new Set(sel.map(r => r[pc])).size };
}
function renderOut() {
  $('outList').innerHTML = Object.entries(FORM_DEFS).map(([code, d]) => {
    const c = formRows(code);
    return `<div class="outitem"><b>${esc(d.label)}</b>
      <span class="sub">${c.panels}盤・${c.n}行　書式: ${tpls[code] ? esc(tpls[code].name) : '<span style="color:var(--warn)">未設定（①設定で選んでください）</span>'}</span>
      <button class="btn" data-make="${code}" ${tpls[code] && c.n ? '' : 'disabled'}>作成して保存</button></div>`;
  }).join('');
}
async function buildForm(code) {
  const t = await kvGet('tpl:' + code);
  if (!t) throw new Error(`${code} の書式が未設定です`);
  return makeForm(code, t.buf, state.mid, state.common);
}
$('outList').onclick = async e => {
  const b = e.target.closest('[data-make]'); if (!b) return;
  b.disabled = true; $('outStatus').textContent = `${b.dataset.make} を作成中…`; $('outStatus').className = 'status';
  try {
    const blob = await buildForm(b.dataset.make);
    if (blob) download(blob, FORM_DEFS[b.dataset.make].out);
    $('outStatus').textContent = `${b.dataset.make} を作成しました。`; $('outStatus').className = 'status ok';
  } catch (err) { console.error(err); $('outStatus').textContent = `作成できませんでした: ${err.message || err}`; $('outStatus').className = 'status err'; }
  b.disabled = false;
};
$('outZip').onclick = async () => {
  const zip = new JSZip(); let n = 0;
  $('outZip').disabled = true;
  try {
    for (const code of Object.keys(FORM_DEFS)) {
      if (!tpls[code] || !formRows(code).n) continue;
      $('outStatus').textContent = `${code} を作成中…`; $('outStatus').className = 'status';
      const blob = await buildForm(code);
      if (blob) { zip.file(FORM_DEFS[code].out, blob); n++; }
    }
    zip.file('中間リスト（添削用）.xlsx', await writeMidXlsx(state.mid, state.common, state.pdfs.map(p => p.name).join(' / ')));
    download(await zip.generateAsync({ type: 'blob' }), '盤図リスト_B02-B08.zip');
    $('outStatus').textContent = `${n}種類の書式と中間リストをZIPにまとめました。`; $('outStatus').className = 'status ok';
  } catch (err) { console.error(err); $('outStatus').textContent = `作成できませんでした: ${err.message || err}`; $('outStatus').className = 'status err'; }
  $('outZip').disabled = false;
};

// ------------------------------------------------------------------ 起動
await load();
{
  const local = await hasSample();
  if (local) $('loadSample').classList.remove('hidden');
  const canTest = local || !!(await kvGet('sample:results'));
  $('testModeRow').classList.toggle('hidden', !canTest);
  if (!canTest) state.settings.testMode = false;
}
initSettings();
markTest();
renderPdfList();
showStep(state.pdfs.length ? 'pdf' : 'settings');
if (!keyStore.get()) showStep('settings');
