// 端末内保存（IndexedDB）。データはこのブラウザの中だけに保存され、外部には送信しない。
const DB = 'banzu-list', STORE = 'kv';
let dbp = null;
function db() {
  if (!dbp) dbp = new Promise((res, rej) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(STORE);
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
  return dbp;
}
async function tx(mode, fn) {
  const d = await db();
  return new Promise((res, rej) => {
    const t = d.transaction(STORE, mode);
    const st = t.objectStore(STORE);
    const req = fn(st);
    t.oncomplete = () => res(req && req.result);
    t.onerror = () => rej(t.error);
    t.onabort = () => rej(t.error);
  });
}
export const kvGet = key => tx('readonly', st => st.get(key)).catch(() => undefined);
export const kvSet = (key, val) => tx('readwrite', st => st.put(val, key));
export const kvDel = key => tx('readwrite', st => st.delete(key));
export const kvKeys = () => tx('readonly', st => st.getAllKeys()).catch(() => []);

// APIキーは「この端末に保存」を選んだときだけ localStorage に置く
export const keyStore = {
  get() { try { return localStorage.getItem('banzu.apiKey') || sessionStorage.getItem('banzu.apiKey') || ''; } catch { return ''; } },
  set(v, remember) {
    try {
      localStorage.removeItem('banzu.apiKey'); sessionStorage.removeItem('banzu.apiKey');
      if (v) (remember ? localStorage : sessionStorage).setItem('banzu.apiKey', v);
    } catch { /* private mode */ }
  },
  remembered() { try { return !!localStorage.getItem('banzu.apiKey'); } catch { return false; } },
};
