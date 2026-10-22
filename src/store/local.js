// Local and OneDrive-synced folders (Chrome and Edge, File System Access API).
// The app is given access to the folder that holds the client timesheet, so it can
// read and write the timesheet, keep team logs in a "Timesheet Logs" subfolder,
// and spot OneDrive conflict copies or a file that is open in Excel.
// Other browsers fall back to opening a single file and downloading updated copies.
import { fingerprint } from '../core/xlsxPatch.js';
import { LOG_FOLDER, isLogFor, parseLog, emptyLog, serialiseLog, ownerOf, freeLogName } from '../core/team.js';
import { StoreError, downloadBytes, saveBytesAs, uid, versionedName } from '../core/util.js';
import * as idb from './idb.js';

export const kind = 'local';
export const label = 'Folder on this computer or OneDrive';
export const canUseFolders = typeof window !== 'undefined' && 'showDirectoryPicker' in window;
// Microsoft 365 only.
export const configured = true;
export const signIn = null;
export const signOut = null;
export const getAccount = () => null;
export const restoreAccount = null;
export const openLink = null;

// Your name (what colleagues see) and a fixed id that links you to your own log files.
const ME_KEY = 'bbtt.me';
const ID_KEY = 'bbtt.meId';
export function getMe() { try { return localStorage.getItem(ME_KEY) || ''; } catch { return ''; } }
export function setMe(name) { try { localStorage.setItem(ME_KEY, name.trim()); } catch { /* ignore */ } }
export function getMeId() {
  try {
    let id = localStorage.getItem(ID_KEY);
    if (!id) { id = uid(); localStorage.setItem(ID_KEY, id); }
    return id;
  } catch { return 'this-browser'; }
}
// Use a log made on another computer as your own (same person, different browser).
export const canClaim = true;
export function claimOwner(owner) { try { localStorage.setItem(ID_KEY, owner); } catch { /* ignore */ } }

export const listRecent = () => idb.listRecent();
export const removeRecent = (id) => idb.removeRecent(id);

async function ensure(handle) {
  const opts = { mode: 'readwrite' };
  if ((await handle.queryPermission(opts)) === 'granted') return true;
  if ((await handle.requestPermission(opts)) === 'granted') return true;
  throw new StoreError('permission', 'Permission to edit files in this folder was not granted.');
}

export async function pickFolder() {
  const dir = await window.showDirectoryPicker({ id: 'bb-timesheets', mode: 'readwrite' });
  await ensure(dir);
  return dir;
}

/** Excel files in the folder that could be timesheets (not lock files, backups or conflict copies). */
export async function listTimesheets(dir) {
  const names = [];
  for await (const [name, h] of dir.entries()) {
    if (h.kind === 'file' && /\.xlsx$/i.test(name) && !name.startsWith('~$') && !/ - Backup - v\d{8}/i.test(name)) names.push(name);
  }
  return names.sort((a, b) => a.localeCompare(b));
}

async function readJsonFiles(logsDir, fileName) {
  const out = [];
  if (!logsDir) return out;
  for await (const [name, h] of logsDir.entries()) {
    if (h.kind !== 'file' || !isLogFor(fileName, name)) continue;
    const f = await h.getFile();
    const log = parseLog(await f.text(), name);
    if (log && (!log.timesheet || log.timesheet.toLowerCase() === fileName.toLowerCase())) out.push({ ...log, fileName: name, modified: f.lastModified });
  }
  return out;
}

export async function openRecent(rec) {
  await ensure(rec.dir);
  return openFromFolder(rec.dir, rec.fileName, rec.id);
}

export async function openFromFolder(dir, fileName, recentId) {
  await ensure(dir);
  // One recent entry per folder and file.
  let id = recentId;
  if (!id) {
    for (const r of await idb.listRecent()) {
      try { if (r.kind === 'local' && r.fileName === fileName && (await r.dir.isSameEntry(dir))) { id = r.id; break; } } catch { /* ignore */ }
    }
  }
  id = id || uid();
  const fileHandle = () => dir.getFileHandle(fileName);
  const logsDir = (create = false) => dir.getDirectoryHandle(LOG_FOLDER, { create }).catch((e) => { if (create) throw e; return null; });

  const source = {
    id,
    name: fileName,
    where: dir.name,
    async remember(client) { await idb.putRecent({ id, kind: 'local', dir, fileName, client }); },

    async read() {
      const f = await (await fileHandle()).getFile();
      const bytes = await f.arrayBuffer();
      return { bytes, version: await fingerprint(bytes), modifiedAt: f.lastModified };
    },

    // Saves only if the file is still the version that was read, otherwise reports a conflict
    // so the change can be re-applied to the newer file.
    async write(bytes, version) {
      await ensure(dir);
      const fh = await fileHandle();
      if (version) {
        const current = await (await fh.getFile()).arrayBuffer();
        if ((await fingerprint(current)) !== version) throw new StoreError('conflict', 'The timesheet changed while saving.');
      }
      try {
        const w = await fh.createWritable();
        await w.write(bytes);
        await w.close();
      } catch (e) {
        throw new StoreError('locked', `Could not save to ${fileName}. If it is open in Excel, close it and try again. Nothing was changed. (${e.message})`);
      }
      return fingerprint(bytes);
    },

    // Cheap check for changes by anyone: timesheet size and time, plus each log file's time.
    async stamp() {
      const f = await (await fileHandle()).getFile();
      const parts = [`${f.lastModified}:${f.size}`];
      const ld = await logsDir();
      if (ld) for await (const [name, h] of ld.entries()) if (h.kind === 'file' && isLogFor(fileName, name)) parts.push(`${name}:${(await h.getFile()).lastModified}`);
      return parts.sort().join('|');
    },

    async warnings() {
      const out = [];
      const base = fileName.replace(/\.xlsx$/i, '');
      for await (const [name, h] of dir.entries()) {
        if (h.kind !== 'file') continue;
        if (name === `~$${fileName}` || name === `~$${fileName.slice(2)}`) out.push('This timesheet is open in Excel on this computer. Close it before saving from the app.');
        else if (name !== fileName && name.toLowerCase().startsWith(`${base.toLowerCase()}-`) && /\.xlsx$/i.test(name)) {
          out.push(`Possible OneDrive conflict copy found: "${name}". Someone's rows may be in that copy. Check it in Excel, and use the Team log to add anything missing.`);
        }
      }
      return out;
    },

    team: {
      available: true,
      identity: () => ({ owner: getMeId(), name: getMe() }),
      async list() { return readJsonFiles(await logsDir(), fileName); },
      async readMine() {
        const mine = (await readJsonFiles(await logsDir(), fileName)).find((l) => ownerOf(l) === getMeId());
        return mine || emptyLog(getMe(), fileName, getMeId());
      },
      // Writes to the file this person already owns, or a new file named after them.
      async writeMine(log) {
        await ensure(dir);
        const ld = await logsDir(true);
        const all = await readJsonFiles(ld, fileName);
        const existing = all.find((l) => ownerOf(l) === getMeId());
        const taken = [];
        for await (const [n] of ld.entries()) taken.push(n);
        const target = existing?.fileName || freeLogName(fileName, getMe() || 'Unnamed', taken);
        const fh = await ld.getFileHandle(target, { create: true });
        const w = await fh.createWritable();
        const { fileName: _f, modified: _m, ...clean } = log;
        await w.write(serialiseLog({ ...clean, owner: getMeId(), person: getMe() }));
        await w.close();
      },
    },

    async saveCopy(bytes) { return (await saveBytesAs(bytes, versionedName(fileName, 'Backup'), dir)).name; },

    async createSibling(name, bytes) {
      try { await dir.getFileHandle(name); throw new StoreError('exists', `"${name}" already exists in this folder.`); } catch (e) { if (e.code === 'exists') throw e; }
      const fh = await dir.getFileHandle(name, { create: true });
      const w = await fh.createWritable();
      await w.write(bytes);
      await w.close();
      return openFromFolder(dir, name);
    },
  };
  return source;
}

/** Fallback for browsers without folder access: one file in memory, updates download as new copies. */
export function openFromBytes(name, initial) {
  let bytes = initial;
  let currentName = name;
  return {
    id: `mem:${name}`,
    get name() { return currentName; },
    where: 'Downloads',
    downloadMode: true,
    async remember() {},
    async read() { return { bytes, version: await fingerprint(bytes) }; },
    async write(next) {
      bytes = next;
      currentName = versionedName(name);
      downloadBytes(next, currentName);
      return fingerprint(next);
    },
    async stamp() { return null; },
    async warnings() { return []; },
    team: null,
    async saveCopy(b) { const n = versionedName(name, 'Backup'); downloadBytes(b, n); return n; },
    createSibling: null,
  };
}
