// Small IndexedDB store: recent timesheets and undo history (kept in this browser only).
const DB = 'bb-timesheet';
const STORES = ['files', 'history', 'mine'];
const HISTORY_MAX = 10;

function db() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 4);
    req.onupgradeneeded = () => {
      const d = req.result;
      if (!d.objectStoreNames.contains('files')) d.createObjectStore('files', { keyPath: 'id' });
      if (!d.objectStoreNames.contains('history')) d.createObjectStore('history', { keyPath: 'fileId' });
      if (!d.objectStoreNames.contains('mine')) d.createObjectStore('mine', { keyPath: 'key' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function tx(store, mode, fn) {
  if (!STORES.includes(store)) throw new Error(`Unknown store ${store}`);
  const d = await db();
  return new Promise((resolve, reject) => {
    const t = d.transaction(store, mode);
    const req = fn(t.objectStore(store));
    t.oncomplete = () => resolve(req?.result);
    t.onerror = () => reject(t.error);
  });
}

export async function listRecent() {
  try { return ((await tx('files', 'readonly', (s) => s.getAll())) || []).sort((a, b) => b.lastOpened - a.lastOpened); } catch { return []; }
}
export const putRecent = (entry) => tx('files', 'readwrite', (s) => s.put({ ...entry, lastOpened: Date.now() }));
export const removeRecent = (id) => tx('files', 'readwrite', (s) => s.delete(id));

export async function getHistory(fileId) {
  try { return (await tx('history', 'readonly', (s) => s.get(fileId)))?.items || []; } catch { return []; }
}
export async function pushHistory(fileId, item) {
  const items = [...(await getHistory(fileId)), item].slice(-HISTORY_MAX);
  await tx('history', 'readwrite', (s) => s.put({ fileId, items }));
  return items;
}
export async function popHistory(fileId) {
  const items = (await getHistory(fileId)).slice(0, -1);
  await tx('history', 'readwrite', (s) => s.put({ fileId, items }));
  return items;
}

// Your own section of a shared team log, kept in this browser so it can be restored if lost.
export async function getMine(key) {
  try { return (await tx('mine', 'readonly', (s) => s.get(key)))?.part || null; } catch { return null; }
}
export const putMine = (key, part) => tx('mine', 'readwrite', (s) => s.put({ key, part }));
