// PDF の読み込み・サムネイル・部分描画（pdf.js）
// iPad のメモリ制限に合わせ、ページ全体を高解像度で描かず、必要な範囲だけを小さなキャンバスに描く。
/* global pdfjsLib */
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

export async function openPdf(file) {
  const data = new Uint8Array(await file.arrayBuffer());
  const doc = await pdfjsLib.getDocument({ data }).promise;
  return doc;
}

/** ページの一部（fraction: x0,y0,x1,y1）を scale で描画し canvas を返す */
export async function renderRegion(doc, pageNo, scale, x0 = 0, y0 = 0, x1 = 1, y1 = 1) {
  const page = await doc.getPage(pageNo);
  const vp = page.getViewport({ scale });
  const W = Math.max(1, Math.round((x1 - x0) * vp.width)), H = Math.max(1, Math.round((y1 - y0) * vp.height));
  const canvas = document.createElement('canvas');
  canvas.width = W; canvas.height = H;
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
  await page.render({ canvasContext: ctx, viewport: vp, transform: [1, 0, 0, 1, -x0 * vp.width, -y0 * vp.height] }).promise;
  page.cleanup();
  return canvas;
}

export async function pageSize(doc, pageNo) {
  const page = await doc.getPage(pageNo);
  const vp = page.getViewport({ scale: 1 });
  return { w: vp.width, h: vp.height };
}

/** 長辺 px 指定で描画 */
export async function renderFit(doc, pageNo, longSide, region) {
  const { w, h } = await pageSize(doc, pageNo);
  const [x0, y0, x1, y1] = region || [0, 0, 1, 1];
  const scale = longSide / Math.max((x1 - x0) * w, (y1 - y0) * h);
  return renderRegion(doc, pageNo, scale, x0, y0, x1, y1);
}

export function canvasToB64(canvas, type = 'image/png', q = 0.85) {
  const url = canvas.toDataURL(type, q);
  const b64 = url.slice(url.indexOf(',') + 1);
  canvas.width = canvas.height = 0; // iPad: メモリ解放
  return b64;
}

export async function pageText(doc, pageNo) {
  try {
    const page = await doc.getPage(pageNo);
    const tc = await page.getTextContent();
    page.cleanup();
    return tc.items.map(i => i.str).join(' ').replace(/\s+/g, ' ').trim();
  } catch { return ''; }
}

/**
 * 読み取り用の画像一式: 全体図1枚 + 重なりのあるタイル（各辺 1400px 程度）
 * 返り値: [{label, b64, w, h}]
 */
export async function extractionImages(doc, pageNo) {
  const { w, h } = await pageSize(doc, pageNo);
  const out = [];
  const ov = await renderFit(doc, pageNo, 1400);
  out.push({ label: '全体図', w: ov.width, h: ov.height, b64: canvasToB64(ov) });
  const scale = Math.min(5, Math.max(3, 3600 / Math.max(w, h)));
  const cols = Math.max(1, Math.ceil((w * scale) / 1400)), rows = Math.max(1, Math.ceil((h * scale) / 1400));
  const o = 0.03;
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
    const x0 = Math.max(0, c / cols - o), x1 = Math.min(1, (c + 1) / cols + o);
    const y0 = Math.max(0, r / rows - o), y1 = Math.min(1, (r + 1) / rows + o);
    const cv = await renderRegion(doc, pageNo, scale, x0, y0, x1, y1);
    out.push({ label: `拡大 ${r + 1}段${c + 1}列（左${Math.round(x0 * 100)}〜${Math.round(x1 * 100)}%、上${Math.round(y0 * 100)}〜${Math.round(y1 * 100)}%）`, w: cv.width, h: cv.height, b64: canvasToB64(cv) });
  }
  return out;
}

/** ページ判定用: 縮小全体図 + 右下の図名欄 */
export async function classifyImages(doc, pageNo) {
  const ov = await renderFit(doc, pageNo, 900);
  const tb = await renderFit(doc, pageNo, 900, [0.55, 0.86, 1, 1]);
  return [{ label: '全体', b64: canvasToB64(ov, 'image/jpeg', 0.8) }, { label: '図名欄', b64: canvasToB64(tb, 'image/png') }];
}
