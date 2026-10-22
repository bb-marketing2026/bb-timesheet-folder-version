// Timesheets on this computer or in a OneDrive-synced folder (Chrome and Edge, File System Access API).
// You open the timesheet file itself and the app saves straight back into it.
// For the team log you connect the folder that holds the timesheet once (the folder picker opens
// in the right place and the app checks it is the right folder). Each person then keeps their own
// log file in a "Timesheet Logs" subfolder, which only their app writes, so logs never clash.
// Improvements over a plain shared folder:
//   - your browser keeps a copy of your own log and puts it back if a OneDrive conflict loses it
//   - duplicate copies of one person's log (OneDrive conflict copies) are merged, never doubled
//   - warnings when the timesheet is open in Excel or OneDrive has made a conflict copy
// Other browsers fall back to opening a single file and downloading updated copies.
import { fingerprint } from '../core/xlsxPatch.js';
import { LOG_FOLDER, isLogFor, freeLogName, emptyLog, docToLogs, mergePart, partMissingFrom, logToPart, emptyPart, parsePersonLog, personLogJson } from '../core/team.js';
import { StoreError, downloadBytes, saveBytesAs, uid, versionedName, XLSX_TYPE } from '../core/util.js';
import * as idb from './idb.js';

export const kind = 'local';
export const label = 'File on this computer or OneDrive';
export const canSaveInPlace = typeof window !== 'undefined' && 'showOpenFilePicker' in window && 'showDirectoryPicker' in window;
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

export async function pickFile() {
  const [handle] = await window.showOpenFilePicker({ id: 'bb-timesheets', types: [XLSX_TYPE], excludeAcceptAllOption: true });
  await ensure(handle);
  return handle;
}

async function readText(handle) { return (await handle.getFile()).text(); }
async function writeText(handle, text) {
  const w = await handle.createWritable();
  await w.write(text);
  await w.close();
}

export async function openRecent(rec) {
  if (rec.dir && rec.fileName) {
    // Folder access covers the timesheet inside it, so one permission prompt is enough.
    try {
      await ensure(rec.dir);
      return openFile(await rec.dir.getFileHandle(rec.fileName), { dir: rec.dir, recentId: rec.id });
    } catch (e) { if (!rec.file) throw e; }
  }
  await ensure(rec.file);
  return openFile(rec.file, { recentId: rec.id });
}

/**
 * Open a timesheet from its file handle. `dir` (the folder holding it) is optional: without it the
 * timesheet works as normal and only the team log and folder warnings are unavailable.
 */
export async function openFile(fileHandle, { dir = null, recentId } = {}) {
  await ensure(fileHandle);
  const fileName = fileHandle.name;
  // One recent entry per file; reuse a folder connected earlier if permission is still there.
  let id = recentId;
  let remembered = dir;
  for (const r of await idb.listRecent()) {
    try {
      if (r.kind !== 'local') continue;
      const same = (r.id && r.id === id) || (r.file && (await r.file.isSameEntry(fileHandle)));
      if (same) { id = r.id; remembered = remembered || r.dir || null; break; }
    } catch { /* ignore */ }
  }
  if (remembered && !dir) {
    try { if ((await remembered.queryPermission({ mode: 'readwrite' })) === 'granted') dir = remembered; } catch { /* ignore */ }
  }
  id = id || uid();
  const mineKey = () => `${id}:${getMeId()}`;
  const logsDir = (create = false) => (dir ? dir.getDirectoryHandle(LOG_FOLDER, { create }).catch((e) => { if (create) throw e; return null; }) : Promise.resolve(null));

  // Every log file for this timesheet, merged per person (conflict copies of one log become one).
  async function readAll() {
    const ld = await logsDir();
    const people = {};
    const files = {};
    if (ld) {
      for await (const [name, h] of ld.entries()) {
        if (h.kind !== 'file' || !isLogFor(fileName, name)) continue;
        const parsed = parsePersonLog(await readText(h), name);
        if (!parsed || (parsed.timesheet && parsed.timesheet.toLowerCase() !== fileName.toLowerCase())) continue;
        people[parsed.owner] = mergePart(people[parsed.owner], parsed.part);
        // Remember this person's main file: the shortest name (conflict copies get longer names).
        if (!files[parsed.owner] || name.length < files[parsed.owner].length) files[parsed.owner] = name;
      }
    }
    return { ld, people, files };
  }
  async function myPart(people) {
    return mergePart(people[getMeId()] || emptyPart(), (await idb.getMine(mineKey())) || emptyPart());
  }
  async function writeMyFile(ld, files, part) {
    const taken = [];
    for await (const [n] of ld.entries()) taken.push(n);
    const target = files[getMeId()] || freeLogName(fileName, getMe() || 'Unnamed', taken);
    await writeText(await ld.getFileHandle(target, { create: true }), personLogJson(getMeId(), fileName, part));
  }

  const source = {
    id,
    name: fileName,
    where: dir ? dir.name : '',
    needsFolder: !dir,
    async remember(client) { await idb.putRecent({ id, kind: 'local', file: fileHandle, dir, fileName, client }); },

    async read() {
      const f = await fileHandle.getFile();
      const bytes = await f.arrayBuffer();
      return { bytes, version: await fingerprint(bytes), modifiedAt: f.lastModified };
    },

    // Saves only if the file is still the version that was read, otherwise reports a conflict
    // so the change can be re-applied to the newer file.
    async write(bytes, version) {
      await ensure(fileHandle);
      if (version) {
        const current = await (await fileHandle.getFile()).arrayBuffer();
        if ((await fingerprint(current)) !== version) throw new StoreError('conflict', 'The timesheet changed while saving.');
      }
      try {
        const w = await fileHandle.createWritable();
        await w.write(bytes);
        await w.close();
      } catch (e) {
        throw new StoreError('locked', `Could not save to ${fileName}. If it is open in Excel, close it and try again. Nothing was changed. (${e.message})`);
      }
      return fingerprint(bytes);
    },

    // Cheap check for changes by anyone: timesheet size and time, plus each log file's time.
    async stamp() {
      const f = await fileHandle.getFile();
      const parts = [`${f.lastModified}:${f.size}`];
      const ld = await logsDir();
      if (ld) for await (const [name, h] of ld.entries()) if (h.kind === 'file' && isLogFor(fileName, name)) parts.push(`${name}:${(await h.getFile()).lastModified}`);
      return parts.sort().join('|');
    },

    async warnings() {
      if (!dir) return [];
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

    // Team log: connect the folder that holds this timesheet (once).
    async connectFolder() {
      let picked;
      try {
        picked = await window.showDirectoryPicker({ id: 'bb-timesheets', mode: 'readwrite', startIn: fileHandle });
      } catch (e) {
        if (e.name === 'AbortError') throw e;
        picked = await window.showDirectoryPicker({ id: 'bb-timesheets', mode: 'readwrite' });
      }
      await ensure(picked);
      const path = await picked.resolve(fileHandle);
      if (!path || path.length !== 1) {
        throw new StoreError('input', `"${picked.name}" is not the folder that holds ${fileName}. Choose the folder the timesheet is saved in.`);
      }
      const next = await openFile(fileHandle, { dir: picked, recentId: id });
      await next.remember();
      return next;
    },

    team: dir ? {
      available: true,
      folderLabel: `${dir.name} / ${LOG_FOLDER}`,
      identity: () => ({ owner: getMeId(), name: getMe() }),
      async list() {
        const { ld, people, files } = await readAll();
        // Put back anything of mine my file is missing (lost in a OneDrive conflict, or deleted).
        const mine = await myPart(people);
        if ((mine.entries.length || Object.keys(mine.added).length) && partMissingFrom(mine, people[getMeId()])) {
          try {
            if ((await dir.queryPermission({ mode: 'readwrite' })) === 'granted') {
              const ld2 = ld || (await logsDir(true));
              await writeMyFile(ld2, files, mine);
              people[getMeId()] = mine;
            }
          } catch { /* read-only for now */ }
        }
        return docToLogs({ timesheet: fileName, people }).map((l) => ({ ...l, fileName: files[l.owner] || '' }));
      },
      async readMine() {
        const { people } = await readAll();
        const [log] = docToLogs({ timesheet: fileName, people: { [getMeId()]: await myPart(people) } });
        return log || emptyLog(getMe(), fileName, getMeId());
      },
      async writeMine(log) {
        await ensure(dir);
        const { people, files } = await readAll();
        const part = logToPart(log, await myPart(people), getMe());
        await idb.putMine(mineKey(), part);
        await writeMyFile(await logsDir(true), files, mergePart(people[getMeId()], part));
      },
    } : null,

    async saveCopy(bytes) { return (await saveBytesAs(bytes, versionedName(fileName, 'Backup'), dir || fileHandle)).name; },

    // New client timesheet: saved where you choose (the dialog opens next to this one), then opened.
    async createSibling(name, bytes) {
      const saved = await saveBytesAs(bytes, name, dir || fileHandle);
      if (!saved.handle) return null;
      return openFile(saved.handle);
    },
  };
  return source;
}

/** Fallback for browsers without file access (Firefox, Safari): one file in memory, updates download as new copies. */
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
    needsFolder: false,
    connectFolder: null,
    async saveCopy(b) { const n = versionedName(name, 'Backup'); downloadBytes(b, n); return n; },
    createSibling: null,
  };
}
