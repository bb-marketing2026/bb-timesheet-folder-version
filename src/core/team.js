// Team log: each person keeps their own small JSON file next to the timesheet, in a
// "Timesheet Logs" folder. Only the owner ever writes their file, so two people can
// never overwrite each other's log. Staff names live only in these files, never in
// the timesheet, which keeps the Brainbox template unchanged.
//
// Each log has a fixed owner id, so a person can change the name colleagues see without
// creating a second log. The file keeps the name it was first created with.
//
// A log holds:
//   owner    fixed id (Microsoft 365: the sign-in account; folder edition: an id kept in the browser)
//   person   the name colleagues see, which the person can change at any time
//   entries  rows this person tracked: { id, date, tag, activity, hours, sessions, savedAt }
//   added    entries this person put into the timesheet (theirs or a colleague's): { entryId, row, at }

export const LOG_FOLDER = 'Timesheet Logs';
const APP = 'brainbox-timesheet-tracker';

const round2 = (n) => Math.round(n * 100) / 100;
const baseName = (timesheetName) => timesheetName.replace(/\.xlsx$/i, '');
const safePart = (s) => String(s).replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim();

export function logFileName(timesheetName, person) {
  return `${baseName(timesheetName)} - ${safePart(person)}.json`;
}

/** Log files that belong to this timesheet, from a list of file names in the logs folder. */
export function isLogFor(timesheetName, fileName) {
  const prefix = `${baseName(timesheetName)} - `;
  return fileName.startsWith(prefix) && fileName.toLowerCase().endsWith('.json');
}

export function emptyLog(person, timesheetName, owner) {
  return { app: APP, version: 1, owner, person, timesheet: timesheetName, entries: [], added: [] };
}

/** Owner id of a log; older logs without one are identified by their file name. */
export const ownerOf = (log) => log.owner || `file:${log.fileName}`;

/** A free file name for a new log, never taking a colleague's. */
export function freeLogName(timesheetName, person, taken) {
  const used = new Set([...taken].map((n) => n.toLowerCase()));
  let n = logFileName(timesheetName, person);
  for (let i = 2; used.has(n.toLowerCase()); i++) n = logFileName(timesheetName, `${person} ${i}`);
  return n;
}

/** Other people's logs already using this display name. */
export function nameTakenBy(logs, name, myOwner) {
  const k = name.trim().toLowerCase();
  return logs.filter((l) => ownerOf(l) !== myOwner && (l.person || '').trim().toLowerCase() === k);
}

export function parseLog(text, fallbackPerson) {
  try {
    const j = JSON.parse(text);
    if (j.app !== APP || !Array.isArray(j.entries)) return null;
    return { ...j, person: j.person || fallbackPerson, added: Array.isArray(j.added) ? j.added : [] };
  } catch { return null; }
}

export const serialiseLog = (log) => JSON.stringify(log, null, 2);

export function addEntries(log, entries) {
  const ids = new Set(log.entries.map((e) => e.id));
  return { ...log, entries: [...log.entries, ...entries.filter((e) => !ids.has(e.id))] };
}

export function addMarkers(log, markers) {
  const fresh = new Map(markers.map((m) => [m.entryId, m]));
  return { ...log, added: [...log.added.filter((m) => !fresh.has(m.entryId)), ...fresh.values()] };
}

export function removeMarkers(log, entryIds) {
  const drop = new Set(entryIds);
  return { ...log, added: log.added.filter((m) => !drop.has(m.entryId)) };
}

const rowKey = (date, tag, activity, hours) => `${date}|${tag.trim().toLowerCase()}|${activity.trim().toLowerCase()}|${round2(hours).toFixed(2)}`;

/**
 * Everyone's entries with a status:
 *   pending  not in the timesheet yet
 *   added    put into the timesheet, and a matching row is there
 *   check    marked as added, but no matching row was found (edited or deleted in Excel,
 *            or lost in a sync conflict). Shown so someone can look, never re-added automatically.
 */
// ---------- Notes on entries: archive, delete, removed from timesheet ----------
// Anyone can archive or delete any entry, or record that its row was deleted from the timesheet.
// Each note lives in the log of the person who made it, as { [entryId]: { [flag]: { v, at, by } } },
// and the latest note for each flag wins across everyone's logs.
export const NOTE_FLAGS = ['archived', 'deleted', 'sheetDeleted'];

export function mergeNotes(a = {}, b = {}) {
  const out = { ...a };
  for (const [id, flags] of Object.entries(b || {})) {
    const cur = { ...(out[id] || {}) };
    for (const [f, n] of Object.entries(flags || {})) if (!cur[f] || (n.at || '') > (cur[f].at || '')) cur[f] = n;
    out[id] = cur;
  }
  return out;
}

/** Set a note on entries in a log (returns a new log). */
export function setNotes(log, entryIds, flag, value, by, now = new Date().toISOString()) {
  const notes = { ...(log.notes || {}) };
  for (const id of entryIds) notes[id] = { ...(notes[id] || {}), [flag]: { v: !!value, at: now, by } };
  return { ...log, notes };
}

/**
 * Everyone's entries with a status:
 *   pending   not in the timesheet yet
 *   added     in the timesheet (matched to a row, see e.row)
 *   removed   was added, but the row is no longer in the timesheet. e.removedBy/removedAt say who
 *             deleted it in the tracker; without them the row was deleted or edited in Excel, or
 *             lost in a sync conflict. Never re-added automatically.
 * Archived entries keep their status and get e.archived = true. Deleted entries are left out.
 */
export function teamView(logs, info) {
  const marks = new Map();
  for (const log of logs) for (const m of log.added) {
    const cur = marks.get(m.entryId);
    if (!cur || (m.at || '') > (cur.at || '')) marks.set(m.entryId, { ...m, by: log.person });
  }
  let notes = {};
  for (const log of logs) notes = mergeNotes(notes, log.notes);
  const on = (id, f) => !!notes[id]?.[f]?.v;
  const pool = new Map();
  for (const e of info?.entries || []) {
    const k = rowKey(e.date, e.tag, e.activity, e.hours);
    if (!pool.has(k)) pool.set(k, []);
    pool.get(k).push(e.row);
  }
  const all = logs.flatMap((log) => log.entries.map((e) => ({ ...e, person: log.person, owner: ownerOf(log) })))
    .filter((e) => !on(e.id, 'deleted'));
  all.sort((a, b) => (a.date === b.date ? (a.savedAt || '').localeCompare(b.savedAt || '') : a.date.localeCompare(b.date)));
  // Match added entries to timesheet rows one to one.
  for (const e of all) {
    e.archived = on(e.id, 'archived');
    const m = marks.get(e.id);
    if (!m) { e.status = 'pending'; continue; }
    e.addedBy = m.by;
    e.addedAt = m.at;
    const rows = pool.get(rowKey(e.date, e.tag, e.activity, e.hours));
    if (rows && rows.length) { e.status = 'added'; e.row = rows.shift(); continue; }
    e.status = 'removed';
    const sd = notes[e.id]?.sheetDeleted;
    if (sd?.v && (sd.at || '') >= (m.at || '')) { e.removedBy = sd.by; e.removedAt = sd.at; }
  }
  return all;
}

/** Turn finished sessions into log entries (one per timesheet row). */
export function sessionsToEntries(rows, sessions, uid) {
  return rows.map((r) => ({
    id: uid(),
    date: r.date,
    tag: r.tag,
    activity: r.activity,
    hours: round2(r.ms / 3600000),
    sessions: sessions
      .filter((s) => s.date === r.date && s.tag.trim().toLowerCase() === r.tag.trim().toLowerCase())
      .map((s) => ({ start: s.start, end: s.end })),
    savedAt: new Date().toISOString(),
  }));
}

// ---------- Per-person log files with safe merging ----------
// Each person's file holds one "section" (part). Markers are kept with a time and a removed flag,
// so an undone marker never comes back when two copies of the same log are merged (for example
// a OneDrive conflict copy, or the copy each browser keeps of its own log).
//
// { app, version: 2, kind: 'person-log', owner, timesheet, person, nameAt,
//   entries: [...], added: { [entryId]: { row, at, removed? } } }

/** Read one person's log file (version 1 or 2) as { owner, part }. */
export function parsePersonLog(text, fileName) {
  let j;
  try { j = JSON.parse(text); } catch { return null; }
  if (!j || j.app !== APP || !Array.isArray(j.entries)) return null;
  const added = Array.isArray(j.added)
    ? Object.fromEntries(j.added.map((m) => [m.entryId, { row: m.row, at: m.at || '' }]))
    : (j.added || {});
  return {
    owner: j.owner || `file:${fileName}`,
    timesheet: j.timesheet || '',
    part: { person: j.person || '', nameAt: j.nameAt || '', entries: j.entries, added, notes: j.notes || {} },
  };
}

export function personLogJson(owner, timesheetName, part) {
  return JSON.stringify({ app: APP, version: 2, kind: 'person-log', owner, timesheet: timesheetName, ...part }, null, 2);
}

export const emptyPart = () => ({ person: '', nameAt: '', entries: [], added: {}, notes: {} });

/** Combine two versions of one person's section: keep every entry, the latest state of each marker and name. */
export function mergePart(a = emptyPart(), b = emptyPart()) {
  const entries = new Map();
  for (const e of [...(a.entries || []), ...(b.entries || [])]) if (!entries.has(e.id)) entries.set(e.id, e);
  const added = { ...(a.added || {}) };
  for (const [id, m] of Object.entries(b.added || {})) if (!added[id] || (m.at || '') > (added[id].at || '')) added[id] = m;
  const useB = (b.nameAt || '') >= (a.nameAt || '');
  return { person: (useB ? b.person : a.person) || a.person || b.person || '', nameAt: useB ? b.nameAt : a.nameAt, entries: [...entries.values()], added, notes: mergeNotes(a.notes, b.notes) };
}

/** True when `part` has something the file's copy is missing (so the file needs putting right). */
export function partMissingFrom(part, inFile = emptyPart()) {
  const ids = new Set((inFile.entries || []).map((e) => e.id));
  if ((part.entries || []).some((e) => !ids.has(e.id))) return true;
  for (const [id, m] of Object.entries(part.added || {})) {
    const f = inFile.added?.[id];
    if (!f || (m.at || '') > (f.at || '')) return true;
  }
  for (const [id, flags] of Object.entries(part.notes || {})) {
    for (const [f, n] of Object.entries(flags)) if ((n.at || '') > (inFile.notes?.[id]?.[f]?.at || '')) return true;
  }
  return (part.nameAt || '') > (inFile.nameAt || '');
}

/** Sections ({ people: { owner: part } }) as a list of per-person logs, the shape the Team log panel uses. */
export function docToLogs(doc) {
  return Object.entries(doc?.people || {}).map(([owner, p]) => ({
    app: APP, owner, person: p.person || 'Unnamed', timesheet: doc.timesheet,
    entries: p.entries || [],
    added: Object.entries(p.added || {}).filter(([, m]) => !m.removed).map(([entryId, m]) => ({ entryId, row: m.row, at: m.at })),
    notes: p.notes || {},
  }));
}

/** Turn an edited log (entries + current markers) back into a section, recording removed markers. */
export function logToPart(log, prev = emptyPart(), name, now = new Date().toISOString()) {
  const added = { ...(prev.added || {}) };
  const current = new Set(log.added.map((m) => m.entryId));
  for (const m of log.added) if (!added[m.entryId] || added[m.entryId].removed || (m.at && m.at > (added[m.entryId].at || ''))) added[m.entryId] = { row: m.row, at: m.at || now };
  for (const id of Object.keys(added)) if (!current.has(id) && !added[id].removed) added[id] = { ...added[id], removed: true, at: now };
  const renamed = name && name !== prev.person;
  return { person: name || prev.person, nameAt: renamed ? now : prev.nameAt || now, entries: log.entries, added, notes: mergeNotes(prev.notes, log.notes) };
}
