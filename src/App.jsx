import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { readWorkbook, analyse, writeWorkbook, createTimesheet, hoursBetween, sessionsToRows } from './core/timesheet.js';
import { placeEntriesSafe, deleteEntrySafe, insertRowsSafe, fingerprint } from './core/xlsxPatch.js';
import { addEntries, addMarkers, removeMarkers, setNotes, teamView, sessionsToEntries, ownerOf, nameTakenBy } from './core/team.js';
import {
  clock, fmtHours, fmtDate, fmtStamp, localDate, localTime, perthIso, longDate, uid, versionedName, saveBytesAs,
} from './core/util.js';
import * as store from './store/index.js';
import { getHistory, pushHistory, popHistory } from './store/idb.js';
import { NewTimesheet, DateField, TimeField, Guide, Manual } from './components.jsx';
import { detectClientTag, genericTags, withClient } from './core/tags.js';

const POLL_MS = 60000; // how often to look for changes made by others
const MAX_RETRIES = 5; // re-apply a change this many times if someone else saved first

function usePersisted(key, initial) {
  const [value, setValue] = useState(() => {
    try { const raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : initial; } catch { return initial; }
  });
  useEffect(() => { try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* private mode */ } }, [key, value]);
  return [value, setValue];
}

function useNow(active) {
  const [now, setNow] = useState(Date.now());
  useEffect(() => {
    if (!active) return undefined;
    const t = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(t);
  }, [active]);
  return now;
}

const STATUS_LABEL = { pending: 'Not added yet', added: 'In timesheet', removed: 'Removed from timesheet' };

export default function App() {
  const [source, setSource] = useState(null); // where the timesheet lives (see store/)
  const [file, setFile] = useState(null); // { bytes, info, sheet }
  const [recent, setRecent] = useState([]);
  const [account, setAccount] = useState(() => store.getAccount());
  const [link, setLink] = useState('');
  const [me, setMeState] = useState(() => store.getMe());
  const [nameEdit, setNameEdit] = useState(false);
  const [nameDraft, setNameDraft] = useState('');
  const [nameConflict, setNameConflict] = useState(null); // { name, owner } when a log already uses the name
  const [clientTag, setClientTag] = useState(''); // client short name for the open timesheet, e.g. "RAWCS"
  const [clientAuto, setClientAuto] = useState(false); // true while it is the guess from the file, not yet saved
  const [clientEdit, setClientEdit] = useState(false);
  const [clientDraft, setClientDraft] = useState('');
  const [timer, setTimer] = usePersisted('bbtt.timer', { tag: '', activity: '', start: null, end: null });
  const [sessions, setSessions] = usePersisted('bbtt.sessions', []);
  const [combine, setCombine] = usePersisted('bbtt.combine', true);
  const [status, setStatus] = useState(null);
  const [busy, setBusy] = useState(false);
  const [history, setHistory] = useState([]);
  const [confirm, setConfirm] = useState(null);
  const [shown, setShown] = useState(8);
  const [insertCount, setInsertCount] = useState('50');
  const [showNew, setShowNew] = useState(false);
  const [showManual, setShowManual] = useState(false);
  // How to use: open on the first visit, then as you left it.
  const [guideOpen, setGuideOpen] = useState(() => { try { return localStorage.getItem('bbtt.guideOpen') !== '0'; } catch { return true; } });
  const toggleGuide = (v) => { setGuideOpen(v); try { localStorage.setItem('bbtt.guideOpen', v ? '1' : '0'); } catch { /* ignore */ } };
  const [logs, setLogs] = useState([]);
  const [teamError, setTeamError] = useState('');
  const [teamFilter, setTeamFilter] = useState({ person: 'all', status: 'open' });
  const [selected, setSelected] = useState(() => new Set());
  const [warnings, setWarnings] = useState([]);
  const [checked, setChecked] = useState(null); // { at, stamp }
  const [updates, setUpdates] = useState(false);
  const inputRef = useRef(null);

  const running = !!timer.start && !timer.end;
  const now = useNow(running);
  const team = source?.team || null;
  const myName = me;
  const myOwner = store.kind === 'local' || account ? store.getMeId() : '';

  useEffect(() => { store.listRecent().then(setRecent); }, []);
  useEffect(() => { store.restoreAccount?.().then((a) => a && setAccount(a)).catch(() => {}); }, []);
  useEffect(() => { setMeState(store.getMe()); }, [account]);
  useEffect(() => {
    document.title = running ? `${clock(now - new Date(timer.start))} · ${timer.tag || 'Timer'}` : 'Brainbox Timesheet Tracker';
  }, [running, now, timer.start, timer.tag]);

  const flash = (type, text, opts = {}) => setStatus({ type, text, ...opts }); // opts.undo: offer Undo for a timesheet change
  const fileSessions = useMemo(() => sessions.filter((s) => !s.fileId || (source && s.fileId === source.id)), [sessions, source]);
  const liveHours = running ? (now - new Date(timer.start)) / 3600000 : 0;
  const pendingHours = fileSessions.reduce((sum, s) => sum + hoursBetween(s.start, s.end), 0);
  const currentHours = timer.start ? (timer.end ? hoursBetween(timer.start, timer.end) : liveHours) : 0;
  const info = file?.info;
  const tags = useMemo(() => genericTags(info?.tags || [], clientTag), [info, clientTag]); // generic activity tags
  const fullTag = (t) => withClient(t, clientTag);

  // ---------- loading ----------
  const loadTeam = useCallback(async (src) => {
    if (!src?.team) { setLogs([]); return []; }
    try {
      const list = await src.team.list();
      setLogs(list);
      setTeamError('');
      return list;
    } catch (e) {
      setTeamError(`Team log not available: ${e.message}`);
      setLogs([]);
      return [];
    }
  }, []);

  async function show(src, bytes, preferredSheet) {
    const info2 = analyse(await readWorkbook(bytes), preferredSheet);
    setFile({ bytes, info: info2, sheet: info2.sheetName });
    // Client tag: saved for this timesheet, otherwise the guess from the file.
    const detected = detectClientTag(info2);
    let saved = null;
    try { saved = localStorage.getItem(`bbtt.clientTag.${src.id}`); } catch { /* ignore */ }
    const ct = saved !== null ? saved : detected;
    setClientTag(ct);
    setClientAuto(saved === null);
    if (!timer.tag && info2.tags[0]) setTimer((t) => ({ ...t, tag: genericTags(info2.tags, ct)[0] || '' }));
    return info2;
  }

  async function refreshAll(src = source, { quiet = false, preferredSheet } = {}) {
    if (!src) return null;
    const before = { rows: file?.info?.entries.length ?? null, logs: new Map(logs.map((l) => [ownerOf(l), l.entries.length])) };
    const { bytes } = await src.read();
    const info2 = await show(src, bytes, preferredSheet || file?.sheet);
    const list = await loadTeam(src);
    setWarnings(await src.warnings().catch(() => []));
    const stamp = await src.stamp().catch(() => null);
    setChecked({ at: Date.now(), stamp });
    setUpdates(false);
    if (!quiet && before.rows !== null) {
      const notes = [];
      const diff = info2.entries.length - before.rows;
      if (diff) notes.push(`timesheet has ${Math.abs(diff)} ${diff > 0 ? 'more' : 'fewer'} row${Math.abs(diff) > 1 ? 's' : ''}`);
      for (const l of list) {
        const n = l.entries.length - (before.logs.get(ownerOf(l)) ?? 0);
        if (n > 0 && ownerOf(l) !== store.getMeId()) notes.push(`${n} new entr${n > 1 ? 'ies' : 'y'} from ${l.person}`);
      }
      flash('info', notes.length ? `Updated: ${notes.join(', ')}.` : 'Up to date. No new changes.');
    }
    return info2;
  }

  async function connect(src, message) {
    setBusy(true);
    setStatus(null);
    try {
      setSource(src);
      setConfirm(null);
      setSelected(new Set());
      setClientEdit(false);
      const info2 = await refreshAll(src, { quiet: true });
      await syncName(src);
      await src.remember(info2.client);
      setRecent(await store.listRecent());
      setHistory(await getHistory(src.id));
      flash('info', message || `Opened ${src.name}.`);
    } catch (e) {
      flash('error', e.message);
    } finally { setBusy(false); }
  }

  // Your log carries the name colleagues see. If you changed it, update this timesheet's log.
  async function syncName(src) {
    if (!src?.team || !store.getMe()) return;
    try {
      const list = await src.team.list();
      const mine = list.find((l) => ownerOf(l) === store.getMeId());
      if (mine && mine.person !== store.getMe()) { await src.team.writeMine(mine); await loadTeam(src); }
    } catch { /* team log not available */ }
  }

  function saveClientTag() {
    if (!source) return;
    const v = clientDraft.trim();
    try { localStorage.setItem(`bbtt.clientTag.${source.id}`, v); } catch { /* ignore */ }
    setClientTag(v);
    setClientAuto(false);
    setClientEdit(false);
    flash('ok', v ? `Client tag for ${source.name} is ${v}. "Content Planning" is written as "${v} Content Planning".` : `No client tag for ${source.name}. Activity tags are written exactly as typed.`);
  }

  async function saveName(draft, claim) {
    const name = draft.trim();
    if (!name) return;
    const taken = nameTakenBy(logs, name, store.getMeId());
    if (taken.length && !claim) {
      if (store.canClaim) { setNameConflict({ name, owner: ownerOf(taken[0]) }); return; }
      flash('error', `"${name}" is already used by a colleague in this team log. Add a surname or initial so people can tell you apart.`);
      return;
    }
    if (claim) store.claimOwner(claim);
    store.setMe(name);
    setMeState(name);
    setNameEdit(false);
    setNameConflict(null);
    if (!source?.team) { flash('ok', `You now appear as ${name}.`); return; }
    await guarded(async () => {
      if (logs.some((l) => ownerOf(l) === store.getMeId())) await source.team.writeMine(await source.team.readMine());
      await loadTeam(source);
      flash('ok', claim
        ? `Linked to your existing team log as ${name}.`
        : `You now appear as ${name} in the team log. Colleagues see it after they press Refresh. It is never written into the timesheet.`);
    });
  }

  // Watch for changes by others (cheap check, nothing is reloaded until you press Refresh).
  useEffect(() => {
    if (!source || !checked) return undefined;
    const t = setInterval(async () => {
      try { const s = await source.stamp(); if (s && checked.stamp && s !== checked.stamp) setUpdates(true); } catch { /* offline */ }
    }, POLL_MS);
    return () => clearInterval(t);
  }, [source, checked]);

  // ---------- opening (local) ----------
  async function chooseFile() {
    setStatus(null);
    if (!store.canSaveInPlace) { inputRef.current?.click(); return; }
    try {
      const handle = await store.pickFile();
      const src = await store.openFile(handle);
      await connect(src, src.needsFolder
        ? `Opened ${src.name}. To share hours with colleagues, click "Connect team log" in the Team log panel.`
        : `Opened ${src.name} with its team log.`);
    } catch (e) { if (e.name !== 'AbortError') flash('error', e.message); }
  }

  // Local edition: the team log lives in the folder that holds the timesheet (connected once).
  async function connectTeamFolder() {
    if (!source?.connectFolder) return;
    try {
      const next = await source.connectFolder();
      await connect(next, 'Team log connected. Colleagues who open this timesheet and connect its folder share the same log.');
    } catch (e) { if (e.name !== 'AbortError') flash('error', e.message); }
  }

  async function onFileInput(e) {
    const f = e.target.files?.[0];
    e.target.value = '';
    if (!f) return;
    await connect(store.openFromBytes(f.name, await f.arrayBuffer()),
      `Opened ${f.name}. This browser cannot save in place or share a team log, so each update downloads a new copy. Use Chrome or Edge for the full app.`);
  }

  async function openRecent(r) {
    try { await connect(await store.openRecent(r)); } catch (e) { flash('error', e.message); }
  }

  // ---------- opening (Microsoft 365) ----------
  async function signIn() {
    try { setAccount(await store.signIn()); } catch (e) { if (!/cancel/i.test(e.message)) flash('error', e.message); }
  }
  async function signOut() {
    await store.signOut();
    setAccount(null); setSource(null); setFile(null); setLogs([]);
  }
  async function openFromLink() {
    if (!link.trim()) return;
    setBusy(true);
    try { const src = await store.openLink(link.trim()); setLink(''); setBusy(false); await connect(src); } catch (e) { flash('error', e.message); setBusy(false); }
  }

  // ---------- timer ----------
  function start() { setStatus(null); setTimer((t) => ({ ...t, start: new Date().toISOString(), end: null })); }
  function end() {
    const endIso = new Date().toISOString();
    setSessions((list) => [...list, {
      id: uid(), fileId: source?.id || null, date: localDate(new Date(timer.start)),
      tag: fullTag(timer.tag), activity: timer.activity.trim(), start: timer.start, end: endIso,
    }]);
    setTimer((t) => ({ ...t, end: endIso }));
  }
  const anotherSession = () => setTimer((t) => ({ tag: t.tag, activity: '', start: null, end: null }));
  function addManual() {
    const endD = new Date();
    const startD = new Date(endD.getTime() - 30 * 60000);
    setSessions((list) => [...list, {
      id: uid(), fileId: source?.id || null, date: localDate(startD),
      tag: fullTag(timer.tag), activity: timer.activity.trim(), start: startD.toISOString(), end: endD.toISOString(),
    }]);
  }
  const updateSession = (id, patch) => setSessions((list) => list.map((s) => (s.id === id ? { ...s, ...patch } : s)));
  const setSessionTime = (s, field, hhmm) => hhmm && updateSession(s.id, { [field]: perthIso(s.date, hhmm) });
  const setSessionDate = (s, key) => key && updateSession(s.id, { date: key, start: perthIso(key, localTime(s.start)), end: perthIso(key, localTime(s.end)) });

  // ---------- writing ----------
  // Every change to the timesheet goes through here: read the latest file, apply one small edit,
  // save only if nobody saved in between (otherwise re-apply to their version), keep an undo copy.
  async function commit(change, extra = {}) {
    let lastErr = null;
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      const { bytes: before, version } = await source.read();
      const res = await change(before);
      try {
        await source.write(res.bytes, version);
        const items = await pushHistory(source.id, { before, afterHash: await fingerprint(res.bytes), label: res.label, at: Date.now(), ...extra, ...(res.extra || {}) });
        setHistory(items);
        await show(source, res.bytes, file?.sheet);
        setChecked({ at: Date.now(), stamp: await source.stamp().catch(() => null) });
        return { ...res, retried: attempt - 1 };
      } catch (e) {
        lastErr = e;
        if (e.code !== 'conflict') throw e;
      }
    }
    throw new Error(`Someone else kept saving the timesheet at the same time, so this was not saved after ${MAX_RETRIES} tries. Press Refresh and try again. (${lastErr?.message})`);
  }

  async function guarded(fn) {
    setBusy(true);
    setStatus(null);
    setConfirm(null);
    try { await fn(); } catch (e) {
      flash('error', e.message);
      if (e.rowsNeeded) setInsertCount(String(Math.max(50, e.rowsNeeded)));
      if (/changed in the file/.test(e.message)) { try { await refreshAll(source, { quiet: true }); } catch { /* ignore */ } }
    } finally { setBusy(false); }
  }

  const invalid = fileSessions.filter((s) => !s.tag || !s.activity || hoursBetween(s.start, s.end) <= 0);
  const needsName = !!team && !myName;
  const canSave = !!file && fileSessions.length > 0 && invalid.length === 0 && !running && !busy && !needsName;

  // Puts log entries into the timesheet and records that in my log.
  async function placeEntries(entries) {
    const res = await commit(async (bytes) => {
      const out = await placeEntriesSafe(bytes, file.sheet, entries.map((e) => ({ ...e, ms: e.hours * 3600000 })));
      const hrs = fmtHours(entries.reduce((s, e) => s + e.hours, 0));
      const n = out.rows.length;
      return { ...out, label: `Added ${n} row${n > 1 ? 's' : ''} (${hrs} hrs)`, extra: { entryIds: entries.map((e) => e.id) } };
    });
    if (team) {
      let mine = await team.readMine();
      mine = addMarkers(mine, res.placed.map((p) => ({ entryId: p.id, row: p.row, at: new Date().toISOString() })));
      // Re-adding an entry that was removed or archived makes it current again.
      const reAdded = entries.filter((e) => e.status === 'removed' || e.archived).map((e) => e.id);
      if (reAdded.length) { mine = setNotes(mine, reAdded, 'sheetDeleted', false, myName); mine = setNotes(mine, reAdded, 'archived', false, myName); }
      await team.writeMine(mine);
    }
    return res;
  }

  // Record a note (archived, deleted, sheetDeleted) on team entries in my log. Everyone sees the latest note.
  async function noteEntries(ids, flag, value) {
    if (!team || !ids.length) return;
    await team.writeMine(setNotes(await team.readMine(), ids, flag, value, myName));
  }

  // mode 'add': save to my team log, then into the timesheet. mode 'log': team log only.
  async function saveSessions(mode) {
    if (!canSave) return;
    const list = fileSessions;
    await guarded(async () => {
      const rows = sessionsToRows(list, combine);
      const entries = sessionsToEntries(rows, list, uid);
      const ids = new Set(list.map((s) => s.id));
      const clearSessions = () => {
        setSessions((l) => l.filter((s) => !ids.has(s.id)));
        setTimer((t) => ({ tag: t.tag, activity: '', start: null, end: null }));
      };
      if (team) {
        // 1. My log first: once it is there, the work cannot be lost even if the timesheet save fails.
        const mine = await team.readMine();
        await team.writeMine(addEntries(mine, entries));
        clearSessions();
        if (mode === 'log') {
          await loadTeam(source);
          flash('ok', `Saved ${entries.length} entr${entries.length > 1 ? 'ies' : 'y'} to your team log. Use "Add selected to timesheet" in the Team log when ready.`);
          return;
        }
        try {
          const res = await placeEntries(entries);
          await loadTeam(source);
          flash('ok', `Added ${res.rows.length} row${res.rows.length > 1 ? 's' : ''} to ${source.name}${res.retried ? ` (re-applied after ${res.retried} newer save${res.retried > 1 ? 's' : ''} by others)` : ''}.`, { undo: true });
        } catch (e) {
          await loadTeam(source);
          const err = new Error(`${e.message} Your work is safe in your team log and shows as "Not added yet". Add it from the Team log when the timesheet is free.`);
          err.rowsNeeded = e.rowsNeeded;
          throw err;
        }
      } else {
        const res = await placeEntries(entries);
        clearSessions();
        flash('ok', source.downloadMode
          ? `Added ${res.rows.length} row${res.rows.length > 1 ? 's' : ''}. Downloaded ${source.name}. Replace your original file with it.`
          : `Added ${res.rows.length} row${res.rows.length > 1 ? 's' : ''} to ${source.name}.`, { undo: true });
      }
    });
  }

  const view = useMemo(() => teamView(logs, info), [logs, info]);
  const people = useMemo(() => logs.map((l) => ({ owner: ownerOf(l), name: l.person || 'Unnamed' })).sort((a, b) => a.name.localeCompare(b.name)), [logs]);
  const matchStatus = (e, f) => (f === 'all' ? true : f === 'archived' ? e.archived : !e.archived && (f === 'open' ? e.status !== 'added' : e.status === f));
  const filtered = view.filter((e) => (teamFilter.person === 'all' || e.owner === teamFilter.person) && matchStatus(e, teamFilter.status));
  const picked = filtered.filter((e) => selected.has(e.id));
  const chosen = picked.filter((e) => e.status !== 'added'); // can be added to the timesheet
  const countOf = (f) => view.filter((e) => (teamFilter.person === 'all' || e.owner === teamFilter.person) && matchStatus(e, f)).length;
  const viewingArchived = teamFilter.status === 'archived';

  // Archive (or restore) selected entries: hidden from the list, the timesheet is not touched.
  async function archiveSelected(value) {
    if (!picked.length) return;
    await guarded(async () => {
      await noteEntries(picked.map((e) => e.id), 'archived', value);
      setSelected(new Set());
      await loadTeam(source);
      flash('ok', value
        ? `Archived ${picked.length} entr${picked.length > 1 ? 'ies' : 'y'}. See ${picked.length > 1 ? 'them' : 'it'} under Archived in the status filter.`
        : `Restored ${picked.length} entr${picked.length > 1 ? 'ies' : 'y'} from the archive.`);
    });
  }

  // Delete selected entries from the team log. Rows of entries that are in the timesheet are deleted there too.
  async function deleteSelected() {
    if (!picked.length) return;
    const inSheet = picked.filter((e) => e.status === 'added' && e.row).sort((a, b) => b.row - a.row);
    await guarded(async () => {
      if (inSheet.length) {
        await commit(async (bytes) => {
          let out = bytes;
          // Bottom row first so row numbers above stay valid. Each row is checked against the entry first.
          for (const e of inSheet) {
            out = (await deleteEntrySafe(out, file.sheet, { row: e.row, tag: e.tag, activity: e.activity, hours: e.hours })).bytes;
          }
          return { bytes: out, label: inSheet.length > 1 ? `Deleted ${inSheet.length} rows with their team log entries` : `Deleted row ${inSheet[0].row} with its team log entry`, extra: { deletedIds: picked.map((e) => e.id) } };
        });
      }
      await noteEntries(picked.map((e) => e.id), 'deleted', true);
      setSelected(new Set());
      await loadTeam(source);
      flash('ok', `Deleted ${picked.length} entr${picked.length > 1 ? 'ies' : 'y'} from the team log${inSheet.length ? ` and ${inSheet.length} row${inSheet.length > 1 ? 's' : ''} from the timesheet` : ''}.`, inSheet.length ? { undo: true } : {});
    });
  }

  async function addSelected() {
    if (!chosen.length) return;
    await guarded(async () => {
      const res = await placeEntries(chosen);
      setSelected(new Set());
      await loadTeam(source);
      flash('ok', `Added ${res.rows.length} team entr${res.rows.length > 1 ? 'ies' : 'y'} to ${source.name} in date order${res.retried ? `, re-applied after ${res.retried} newer save${res.retried > 1 ? 's' : ''}` : ''}.`, { undo: true });
    });
  }

  async function deleteRow(entry) {
    // Team log entries that point at this row are marked "Removed from timesheet" (synced on Refresh).
    const linked = view.filter((e) => e.status === 'added' && e.row === entry.row).map((e) => e.id);
    await guarded(async () => {
      await commit(async (bytes) => {
        const { bytes: out } = await deleteEntrySafe(bytes, file.sheet, entry);
        return { bytes: out, label: `Deleted row ${entry.row} (${fmtDate(entry.date)}, ${entry.tag}, ${fmtHours(entry.hours)} hrs)`, extra: { sheetDeletedIds: linked } };
      });
      await noteEntries(linked, 'sheetDeleted', true);
      await loadTeam(source);
      flash('ok', `Deleted row ${entry.row}. Rows below moved up one.${linked.length ? ' Its team log entry now shows "Removed from timesheet" for everyone after they press Refresh.' : ''}`, { undo: true });
    });
  }

  async function insertRows() {
    const n = Number(insertCount);
    await guarded(async () => {
      let r;
      await commit(async (bytes) => { r = await insertRowsSafe(bytes, file.sheet, n); return { bytes: r.bytes, label: `Inserted ${n} blank row${n > 1 ? 's' : ''}` }; });
      flash('ok', `Inserted ${n} blank row${n > 1 ? 's' : ''} above Total Hours, which is now row ${r.totalRow}. Blank rows left: ${r.rowsLeft}.`, { undo: true });
    });
  }

  async function saveCopy() {
    setStatus(null);
    try {
      const { bytes } = await source.read();
      const name = await source.saveCopy(bytes);
      flash('ok', `Backup saved as ${name}. You are still working in ${source.name}.`);
    } catch (e) { if (e.name !== 'AbortError') flash('error', e.message); }
  }

  async function undoLast(force = false) {
    const step = history.at(-1);
    if (!step || busy) return;
    setBusy(true);
    setStatus(null);
    try {
      const { bytes: current, version } = await source.read();
      if (!force && (await fingerprint(current)) !== step.afterHash) { setConfirm({ kind: 'undo' }); return; }
      await source.write(step.before, version);
      if (team && (step.entryIds?.length || step.sheetDeletedIds?.length || step.deletedIds?.length)) {
        let mine = await team.readMine();
        if (step.entryIds?.length) mine = removeMarkers(mine, step.entryIds);
        if (step.sheetDeletedIds?.length) mine = setNotes(mine, step.sheetDeletedIds, 'sheetDeleted', false, myName);
        if (step.deletedIds?.length) mine = setNotes(mine, step.deletedIds, 'deleted', false, myName);
        await team.writeMine(mine);
      }
      setHistory(await popHistory(source.id));
      setConfirm(null);
      await show(source, step.before, file.sheet);
      await loadTeam(source);
      flash('info', `Undone: ${step.label}.${step.entryIds?.length && team ? ' Those entries show as "Not added yet" in the Team log.' : ''}`);
    } catch (e) {
      flash('error', e.code === 'conflict' ? 'The timesheet changed while undoing. Press Refresh and try again.' : e.message);
    } finally { setBusy(false); }
  }

  async function createNew(form) {
    try {
      const bytes = await writeWorkbook(await createTimesheet(form));
      const name = `Brainbox - Timesheet - ${form.client.trim()}.xlsx`;
      if (source?.createSibling) {
        const src = await source.createSibling(name, bytes);
        setShowNew(false);
        if (src) await connect(src, `Created and opened ${src.name}.`);
        else flash('ok', `Downloaded ${name}. Open it to start tracking.`);
      } else {
        const saved = await saveBytesAs(bytes, name);
        setShowNew(false);
        if (saved.handle && store.openFile) await connect(await store.openFile(saved.handle), `Created and opened ${saved.name}.`);
        else flash('ok', `Saved ${saved.name}. ${store.kind === 'local' ? 'Open it to start tracking.' : 'Upload it to SharePoint, copy its link and open it here.'}`);
      }
    } catch (e) { if (e.name !== 'AbortError') flash('error', e.message); }
  }

  const toggle = (id) => setSelected((s) => { const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n; });

  // ---------- render ----------
  const connectPanel = store.kind === 'm365' ? (
    !store.configured ? (
      <p className="warn-box">This app has not been connected to Microsoft 365 yet. Add VITE_MSAL_CLIENT_ID and VITE_MSAL_TENANT_ID in Vercel (see README), then redeploy.</p>
    ) : !account ? (
      <div className="connect">
        <p className="muted">Sign in with your Brainbox Microsoft account to open timesheets stored in SharePoint or OneDrive.</p>
        <button onClick={signIn}>Sign in with Microsoft</button>
      </div>
    ) : (
      <div className="connect">
        <div className="row gap between"><span className="muted">Signed in as <strong>{account.name}</strong></span><button className="link dark" onClick={signOut}>Sign out</button></div>
        <label className="field"><span>Timesheet link (SharePoint or OneDrive, from "Copy link")</span>
          <div className="row gap nowrap">
            <input value={link} placeholder="https://brainbox.sharepoint.com/:x:/…" onChange={(e) => setLink(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && openFromLink()} />
            <button onClick={openFromLink} disabled={busy || !link.trim()}>Open</button>
          </div>
        </label>
      </div>
    )
  ) : null;

  return (
    <>
      <div className="confidential" role="note">
        <strong>Strictly private &amp; confidential</strong>
        <span>Internal Brainbox use only. Do not forward, publish or distribute outside the group.</span>
      </div>
    <div className="app">
      <header className="bar">
        <div className="brand"><span className="dot" />Brainbox Timesheet Tracker{store.kind === 'm365' && <span className="tagline">Microsoft 365</span>}</div>
        <div className="row gap">
          {running && <div className="live"><span className="pulse" />{clock(now - new Date(timer.start))}</div>}
          <button className={`ghost small ${guideOpen ? 'on' : ''}`} aria-expanded={guideOpen} onClick={() => toggleGuide(!guideOpen)}>How to use</button>
          <button className="ghost small" onClick={() => setShowManual(true)}>User manual</button>
        </div>
      </header>
      {guideOpen && <Guide edition={store.kind} onManual={() => setShowManual(true)} onClose={() => toggleGuide(false)} />}

      {/* Your name: shown to colleagues in the Team log, never written into the timesheet. */}
      <div className="namebar">
        <div className="namebar-name">
        {nameEdit || !myName ? (
          <div className="namebar-edit">
            <label htmlFor="me">Your name</label>
            <div className="row gap nowrap">
              <input id="me" value={nameDraft} placeholder="Type your name" maxLength={40} autoComplete="name"
                disabled={store.kind === 'm365' && !account}
                onChange={(e) => { setNameDraft(e.target.value); setNameConflict(null); }}
                onKeyDown={(e) => e.key === 'Enter' && nameDraft.trim() && saveName(nameDraft)} />
              <button className="small" disabled={!nameDraft.trim() || busy || nameDraft.trim() === myName || (store.kind === 'm365' && !account)} onClick={() => saveName(nameDraft)}>Save</button>
              {myName && <button className="ghost small" onClick={() => { setNameEdit(false); setNameConflict(null); }}>Cancel</button>}
            </div>
            <span className="hint">
              {store.kind === 'm365' && !account
                ? 'Sign in with Microsoft first, then save your name.'
                : 'Colleagues see this name next to your entries in the Team log after they press Refresh. It is never written into the timesheet.'}
            </span>
            {nameConflict && (
              <div className="warn-box">
                A team log for "{nameConflict.name}" already exists. Is it yours, from another computer or browser? Only choose Yes if it really is you, otherwise your entries would mix with a colleague's.
                <div className="row gap" style={{ marginTop: 8 }}>
                  <button className="small" onClick={() => saveName(nameConflict.name, nameConflict.owner)}>Yes, it is mine</button>
                  <button className="ghost small" onClick={() => setNameConflict(null)}>No, use another name</button>
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="namebar-show">
            <span>Your name: <strong>{myName}</strong></span>
            <button className="link dark" onClick={() => { setNameDraft(myName); setNameEdit(true); }}>Change</button>
            <span className="hint">Shown to colleagues in the Team log.</span>
          </div>
        )}
        </div>
        <div className="namebar-client">
          {!file ? (
            <div className="namebar-show"><span>Client tag: <span className="muted">open a timesheet</span></span></div>
          ) : clientEdit || (!clientTag && clientAuto && !(info?.tags?.length)) ? (
            <div className="namebar-edit">
              <label htmlFor="client-tag">Client tag</label>
              <div className="row gap nowrap">
                <input id="client-tag" value={clientDraft} placeholder="Client short name" maxLength={30}
                  onChange={(e) => setClientDraft(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && saveClientTag()} />
                <button className="small" onClick={saveClientTag} disabled={busy}>Save</button>
                {clientEdit && <button className="ghost small" onClick={() => setClientEdit(false)}>Cancel</button>}
              </div>
              <span className="hint">Starts every activity tag written for this timesheet, for example "{(clientDraft || 'Client').trim() || 'Client'} Content Planning". Leave empty for none.</span>
            </div>
          ) : (
            <div className="namebar-show">
              <span>Client tag: <strong>{clientTag || 'none'}</strong>{clientAuto && <span className="muted">{clientTag ? ' (from the timesheet)' : ' (not used in this timesheet)'}</span>}</span>
              <button className="link dark" onClick={() => { setClientDraft(clientTag); setClientEdit(true); }}>Change</button>
            </div>
          )}
        </div>
      </div>

      {status && (
        <div className={`status ${status.type}`} role="status">
          <span>{status.text}</span>
          {status.undo && history.length > 0 && <button className="link" onClick={() => undoLast()} disabled={busy}>Undo</button>}
          <button className="close" onClick={() => setStatus(null)} aria-label="Dismiss">×</button>
        </div>
      )}

      <main className="grid">
        <section className="col">
          <div className="card">
            <div className="card-head">
              <h2>Timesheet file</h2>
              <div className="row gap">
                <button className="ghost" onClick={() => setShowNew(true)} disabled={store.kind === 'm365' && !account}>New timesheet</button>
                {store.kind === 'local' && <button onClick={chooseFile} disabled={busy}>{source ? 'Change' : 'Open timesheet'}</button>}
              </div>
            </div>
            <input ref={inputRef} type="file" accept=".xlsx" hidden onChange={onFileInput} />
            {connectPanel}

            {file ? (
              <div className="file">
                <div className="file-name">{source.name}</div>
                <dl className="meta">
                  <dt>Client</dt><dd>{info.client || 'Not set'}</dd>
                  <dt>Project</dt><dd>{info.project || 'Not set'}</dd>
                  {source.where && <><dt>Folder</dt><dd>{source.where}</dd></>}
                  <dt>Sheet</dt>
                  <dd>
                    {info.sheets.length > 1 ? (
                      <select value={file.sheet} onChange={(e) => refreshAll(source, { quiet: true, preferredSheet: e.target.value })}>
                        {info.sheets.map((s) => <option key={s}>{s}</option>)}
                      </select>
                    ) : file.sheet}
                  </dd>
                  <dt>Hours in file</dt><dd><strong>{fmtHours(info.total)}</strong></dd>
                  <dt>Last entry</dt><dd>{fmtDate(info.lastDate) || 'None yet'} (row {info.lastUsedRow})</dd>
                  <dt>Blank rows left</dt>
                  <dd className={info.rowsLeft !== null && info.rowsLeft < 10 ? 'warn' : ''}>
                    {info.rowsLeft ?? 'No Total Hours row found'}{info.rowsLeft !== null && info.rowsLeft < 10 && ' (insert more rows soon)'}
                  </dd>
                  {source.modifiedBy && <><dt>Last saved by</dt><dd>{source.modifiedBy}</dd></>}
                  <dt>Saving</dt><dd>{source.downloadMode ? 'Downloads an updated copy' : 'Straight into this file'}</dd>
                </dl>

                {warnings.map((w) => <p key={w} className="warn-box">{w}</p>)}

                <div className="tools">
                  <div className="tool">
                    <label htmlFor="insert-count">Insert blank rows above Total Hours</label>
                    <div className="row gap nowrap">
                      <input id="insert-count" type="number" min="1" max="500" step="1" inputMode="numeric" className="count"
                        value={insertCount} onChange={(e) => setInsertCount(e.target.value)} />
                      <button className="ghost small" onClick={insertRows}
                        disabled={busy || !info.totalRow || !(Number(insertCount) >= 1 && Number(insertCount) <= 500 && Number.isInteger(Number(insertCount)))}>Insert rows</button>
                    </div>
                    <span className="hint">1 to 500 rows. Formatting is copied and the Total Hours formula is extended.</span>
                  </div>
                  <div className="tool">
                    <label>Backup</label>
                    <div className="row gap"><button className="ghost small" onClick={saveCopy} disabled={busy}>Save a copy</button></div>
                    <span className="hint">Exact copy, named like "{versionedName(source.name, 'Backup')}". You keep working in the original.</span>
                  </div>
                </div>

                {history.length > 0 && (
                  <div className="undo">
                    {confirm?.kind === 'undo' ? (
                      <>
                        <span>The timesheet has changed since this edit, possibly by a colleague. Undoing will also remove those later changes.</span>
                        <div className="row gap">
                          <button className="danger small" onClick={() => undoLast(true)} disabled={busy}>Undo anyway</button>
                          <button className="ghost small" onClick={() => setConfirm(null)}>Keep file</button>
                        </div>
                      </>
                    ) : (
                      <>
                        <span>Your last change: {history.at(-1).label}, {fmtStamp(history.at(-1).at)}</span>
                        <button className="ghost small" onClick={() => undoLast()} disabled={busy}>Undo</button>
                      </>
                    )}
                  </div>
                )}
              </div>
            ) : (
              store.kind === 'local' && <p className="muted">Open a client timesheet (.xlsx). The app reads your activity tags from its Project Description column and saves your hours straight into the file.</p>
            )}

            {recent.length > 0 && (store.kind === 'local' ? store.canSaveInPlace : !!account) && (
              <div className="recent">
                <div className="label">Recent clients</div>
                <div className="chips">
                  {recent.map((r) => (
                    <span key={r.id} className={`chip ${source?.id === r.id ? 'on' : ''}`}>
                      <button className="chip-btn" onClick={() => openRecent(r)} title={r.fileName || r.name}>
                        {r.client ? r.client.split(' - ')[0] : (r.fileName || r.name).replace(/\.xlsx$/i, '')}
                      </button>
                      <button className="chip-x" aria-label={`Remove ${r.fileName || r.name}`} onClick={async () => { await store.removeRecent(r.id); setRecent(await store.listRecent()); }}>×</button>
                    </span>
                  ))}
                </div>
              </div>
            )}
            {store.kind === 'local' && !store.canSaveInPlace && <p className="note">Tip: use Chrome or Edge to save straight into the file and share a team log.</p>}
          </div>

          <div className="card timer">
            <h2>Session</h2>
            <label className="field">
              <span>Activity tag</span>
              <input list="tags" value={timer.tag} placeholder="e.g. Content Planning" onChange={(e) => setTimer((t) => ({ ...t, tag: e.target.value }))} />
              <datalist id="tags">{tags.map((t) => <option key={t} value={t} />)}</datalist>
              <datalist id="fulltags">{(info?.tags || []).map((t) => <option key={t} value={t} />)}</datalist>
              {clientTag && timer.tag.trim() && <span className="tag-preview">Written to the timesheet as <strong>{fullTag(timer.tag)}</strong></span>}
            </label>
            <label className="field">
              <span>Activity</span>
              <textarea rows={2} value={timer.activity} placeholder="What are you working on?" onChange={(e) => setTimer((t) => ({ ...t, activity: e.target.value }))} />
            </label>
            <div className="clock">{clock(timer.start ? (timer.end ? new Date(timer.end) : now) - new Date(timer.start) : 0)}</div>
            <div className="today">{longDate(timer.start ? new Date(timer.start) : new Date())}</div>
            <div className="times">
              <div><span>Start</span><strong>{timer.start ? localTime(timer.start) : '--:--'}</strong></div>
              <div><span>End</span><strong>{timer.end ? localTime(timer.end) : '--:--'}</strong></div>
              <div><span>Hours</span><strong>{fmtHours(currentHours)}</strong></div>
            </div>
            <div className="row gap">
              {!timer.start && <button className="big start" onClick={start} disabled={!timer.tag.trim()}>Start</button>}
              {running && <button className="big stop" onClick={end}>End</button>}
              {timer.end && <button className="big ghost" onClick={anotherSession}>Add another session</button>}
            </div>
            {!timer.tag.trim() && !timer.start && <p className="hint">Pick an activity tag to start.</p>}
            {running && <p className="hint">You can close this tab. The timer keeps counting and picks up when you return.</p>}
          </div>
        </section>

        <section className="col">
          <div className="card">
            <div className="card-head">
              <h2>Sessions to add</h2>
              <button className="ghost small" onClick={addManual} title="Add a 30 minute entry you can edit">+ Manual entry</button>
            </div>
            {fileSessions.length === 0 ? (
              <p className="muted">Ended sessions appear here. Edit the times if you forgot to start or stop.</p>
            ) : (
              <div className="sessions">
                {fileSessions.map((s) => {
                  const h = hoursBetween(s.start, s.end);
                  const bad = !s.tag || !s.activity || h <= 0;
                  return (
                    <div key={s.id} className={`session ${bad ? 'bad' : ''}`}>
                      <div className="s-top">
                        <input list="fulltags" className="s-tag" value={s.tag} placeholder="Activity tag" onChange={(e) => updateSession(s.id, { tag: e.target.value })} />
                        <button className="icon" aria-label="Delete session" onClick={() => setSessions((l) => l.filter((x) => x.id !== s.id))}>×</button>
                      </div>
                      <input className="s-act" value={s.activity} placeholder="Activity (required)" onChange={(e) => updateSession(s.id, { activity: e.target.value })} />
                      <div className="s-times">
                        <DateField value={s.date} onChange={(v) => setSessionDate(s, v)} />
                        <TimeField value={localTime(s.start)} onChange={(v) => setSessionTime(s, 'start', v)} label="Start time" />
                        <span>to</span>
                        <TimeField value={localTime(s.end)} onChange={(v) => setSessionTime(s, 'end', v)} label="End time" />
                        <strong className="s-h">{fmtHours(h)} h</strong>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
            <div className="total"><span>Hours in total</span><strong>{fmtHours(pendingHours + (running ? liveHours : 0))}</strong></div>
            <label className="check">
              <input type="checkbox" checked={combine} onChange={(e) => setCombine(e.target.checked)} />
              Combine sessions with the same tag on the same day into one row
            </label>
            <button className="big primary wide" onClick={() => saveSessions('add')} disabled={!canSave}>{busy ? 'Saving…' : 'Add hours to the file'}</button>
            {team && <button className="ghost wide second" onClick={() => saveSessions('log')} disabled={!canSave}>Save to team log only</button>}
            {!file && fileSessions.length > 0 && <p className="hint">Open a timesheet first.</p>}
            {running && <p className="hint">End the running session before adding hours.</p>}
            {needsName && fileSessions.length > 0 && <p className="hint">Enter your name at the top of the page first.</p>}
            {invalid.length > 0 && <p className="hint">Each session needs a tag, an activity and more than 0 hours.</p>}
          </div>

          {file && (
            <div className="card">
              <div className="card-head">
                <h2>Team log</h2>
                <div className="row gap">
                  {checked && <span className="stamp">Checked {localTime(new Date(checked.at).toISOString())}</span>}
                  <button className={`small ${updates ? 'pulse-btn' : 'ghost'}`} onClick={() => guarded(() => refreshAll(source))} disabled={busy}>
                    {updates ? 'Refresh: new updates' : 'Refresh'}
                  </button>
                </div>
              </div>
              {!team && source.needsFolder ? (
                <div className="connect-team">
                  <p>Share hours with colleagues working on this timesheet. Connect the folder that holds <strong>{source.name}</strong> and the app keeps everyone's log in a "Timesheet Logs" folder beside it.</p>
                  <button className="small" onClick={connectTeamFolder} disabled={busy}>Connect team log</button>
                  <span className="hint">The folder picker opens where the timesheet is saved. Click "Select folder", then "Edit files". Once per timesheet. The app checks you picked the right folder.</span>
                </div>
              ) : !team ? (
                <p className="muted">{source.downloadMode ? 'The team log needs Chrome or Edge.' : 'Team log is not available for this file.'}</p>
              ) : teamError ? (
                <p className="warn-box">{teamError}</p>
              ) : (
                <>
                  <div className="filters">
                    <select aria-label="Person" value={teamFilter.person} onChange={(e) => setTeamFilter((f) => ({ ...f, person: e.target.value }))}>
                      <option value="all">Everyone</option>
                      {people.map((p) => <option key={p.owner} value={p.owner}>{p.name}{p.owner === myOwner ? ' (you)' : ''}</option>)}
                    </select>
                    <select aria-label="Status" value={teamFilter.status} onChange={(e) => setTeamFilter((f) => ({ ...f, status: e.target.value }))}>
                      <option value="open">Needs action ({countOf('open')})</option>
                      <option value="pending">Not added yet ({countOf('pending')})</option>
                      <option value="removed">Removed from timesheet ({countOf('removed')})</option>
                      <option value="added">In timesheet ({countOf('added')})</option>
                      <option value="archived">Archived ({countOf('archived')})</option>
                      <option value="all">All ({countOf('all')})</option>
                    </select>
                  </div>
                  {filtered.length === 0 ? (
                    <p className="muted">{view.length ? 'Nothing here for this filter.' : 'No team entries yet. Entries appear here when anyone saves sessions for this timesheet.'}</p>
                  ) : (
                    <table className="rows team">
                      <thead><tr><th aria-label="Select" /><th>Person</th><th>Date</th><th>Tag / activity</th><th>Hrs</th><th>Status</th></tr></thead>
                      <tbody>
                        {filtered.slice(0, 60).map((e) => (
                          <tr key={e.id} className={selected.has(e.id) ? 'sel' : ''}>
                            <td><input type="checkbox" aria-label={`Select ${e.activity}`} checked={selected.has(e.id)} onChange={() => toggle(e.id)} /></td>
                            <td>{e.person}</td>
                            <td>{fmtDate(e.date)}</td>
                            <td><strong>{e.tag}</strong><br />{e.activity}</td>
                            <td>{fmtHours(e.hours)}</td>
                            <td>
                              <span className={`pill ${e.status}`} title={e.status === 'removed'
                                ? (e.removedBy ? `Deleted from the timesheet by ${e.removedBy}, ${fmtStamp(Date.parse(e.removedAt))}.` : 'The row is no longer in the timesheet. It may have been deleted or edited in Excel, or lost in a sync conflict.')
                                : e.addedBy ? `Added by ${e.addedBy}` : ''}>{STATUS_LABEL[e.status]}</span>
                              {e.status === 'removed' && <span className="pill-note">{e.removedBy ? `by ${e.removedBy}` : 'not found in file'}</span>}
                              {e.archived && <span className="pill archived">Archived</span>}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                  {confirm?.kind === 'team-delete' ? (
                    <div className="confirm team-confirm">
                      <span>
                        Delete {picked.length} entr{picked.length > 1 ? 'ies' : 'y'} from the team log for everyone?
                        {picked.some((e) => e.status === 'added') && ` ${picked.filter((e) => e.status === 'added').length} of them ${picked.filter((e) => e.status === 'added').length > 1 ? 'are' : 'is'} in the timesheet and the row${picked.filter((e) => e.status === 'added').length > 1 ? 's' : ''} will be deleted there too.`}
                        {' '}To keep them, use Archive instead.
                      </span>
                      <div className="row gap">
                        <button className="danger small" onClick={deleteSelected} disabled={busy}>Delete</button>
                        <button className="ghost small" onClick={() => setConfirm(null)}>Cancel</button>
                      </div>
                    </div>
                  ) : (
                    <div className="team-foot">
                      <button className="small" onClick={addSelected} disabled={busy || !chosen.length}>
                        Add selected to timesheet{chosen.length ? ` (${chosen.length}, ${fmtHours(chosen.reduce((s, e) => s + e.hours, 0))} hrs)` : ''}
                      </button>
                      <button className="ghost small" onClick={() => archiveSelected(!viewingArchived)} disabled={busy || !picked.length}>
                        {viewingArchived ? 'Restore selected' : 'Archive selected'}
                      </button>
                      <button className="ghost small danger-text" onClick={() => setConfirm({ kind: 'team-delete' })} disabled={busy || !picked.length}>Delete selected</button>
                      {filtered.some((e) => e.status !== 'added') && (
                        <button className="link dark" onClick={() => setSelected(new Set(filtered.filter((e) => e.status !== 'added').map((e) => e.id)))}>Select all needing action</button>
                      )}
                      {picked.length > 0 && <button className="link dark" onClick={() => setSelected(new Set())}>Clear</button>}
                    </div>
                  )}
                  <p className="hint">Add puts entries into the timesheet in date order. Archive hides entries and leaves the timesheet as it is. Delete removes entries for everyone, and their rows if they are in the timesheet. Press Refresh to see changes made by colleagues. Names are never written into the timesheet.{team.folderLabel ? ` Logs are kept in ${team.folderLabel}.` : ''}</p>
                </>
              )}
            </div>
          )}

          {info && info.entries.length > 0 && (
            <div className="card">
              <h2>Latest rows in the file</h2>
              <table className="rows">
                <thead><tr><th>Date</th><th>Tag</th><th>Activity</th><th>Hrs</th><th aria-label="Delete" /></tr></thead>
                <tbody>
                  {info.entries.slice(-shown).reverse().map((r) => (
                    confirm?.kind === 'delete' && confirm.row === r.row ? (
                      <tr key={r.row} className="confirm-row">
                        <td colSpan={5}>
                          <div className="confirm">
                            <span>Delete row {r.row}: <strong>{r.tag}</strong>, {r.activity} ({fmtHours(r.hours)} hrs)? Rows below move up one.</span>
                            <div className="row gap">
                              <button className="danger small" onClick={() => deleteRow(r)} disabled={busy}>Delete</button>
                              <button className="ghost small" onClick={() => setConfirm(null)}>Cancel</button>
                            </div>
                          </div>
                        </td>
                      </tr>
                    ) : (
                      <tr key={r.row}>
                        <td>{fmtDate(r.date)}</td><td>{r.tag}</td><td>{r.activity}</td><td>{fmtHours(r.hours)}</td>
                        <td className="del">
                          <button className="icon trash" title={`Delete row ${r.row}`} aria-label={`Delete row ${r.row}`} onClick={() => setConfirm({ kind: 'delete', row: r.row })} disabled={busy}>
                            <svg viewBox="0 0 16 16" width="15" height="15" aria-hidden="true"><path d="M2.5 4h11M6 4V2.5h4V4M4 4l.7 9.5h6.6L12 4M6.8 6.5v5M9.2 6.5v5" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" /></svg>
                          </button>
                        </td>
                      </tr>
                    )
                  ))}
                </tbody>
              </table>
              {info.entries.length > shown && <button className="ghost small more" onClick={() => setShown((n) => n + 10)}>Show 10 more</button>}
            </div>
          )}
        </section>
      </main>

      {showNew && <NewTimesheet onCancel={() => setShowNew(false)} onCreate={createNew} />}
      {showManual && <Manual edition={store.kind} onClose={() => setShowManual(false)} />}
    </div>
    </>
  );
}
