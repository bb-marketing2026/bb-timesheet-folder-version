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
  const ids = new Set(log.added.map((m) => m.entryId));
  return { ...log, added: [...log.added, ...markers.filter((m) => !ids.has(m.entryId))] };
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
export function teamView(logs, info) {
  const marks = new Map();
  for (const log of logs) for (const m of log.added) if (!marks.has(m.entryId)) marks.set(m.entryId, { ...m, by: log.person });
  const pool = new Map();
  for (const e of info?.entries || []) {
    const k = rowKey(e.date, e.tag, e.activity, e.hours);
    pool.set(k, (pool.get(k) || 0) + 1);
  }
  const all = logs.flatMap((log) => log.entries.map((e) => ({ ...e, person: log.person, owner: ownerOf(log) })));
  all.sort((a, b) => (a.date === b.date ? (a.savedAt || '').localeCompare(b.savedAt || '') : a.date.localeCompare(b.date)));
  // Match added entries to timesheet rows one to one.
  for (const e of all) {
    const m = marks.get(e.id);
    if (!m) { e.status = 'pending'; continue; }
    e.addedBy = m.by;
    e.addedAt = m.at;
    const k = rowKey(e.date, e.tag, e.activity, e.hours);
    if ((pool.get(k) || 0) > 0) { pool.set(k, pool.get(k) - 1); e.status = 'added'; } else e.status = 'check';
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
    part: { person: j.person || '', nameAt: j.nameAt || '', entries: j.entries, added },
  };
}

export function personLogJson(owner, timesheetName, part) {
  return JSON.stringify({ app: APP, version: 2, kind: 'person-log', owner, timesheet: timesheetName, ...part }, null, 2);
}

export const emptyPart = () => ({ person: '', nameAt: '', entries: [], added: {} });

/** Combine two versions of one person's section: keep every entry, the latest state of each marker and name. */
export function mergePart(a = emptyPart(), b = emptyPart()) {
  const entries = new Map();
  for (const e of [...(a.entries || []), ...(b.entries || [])]) if (!entries.has(e.id)) entries.set(e.id, e);
  const added = { ...(a.added || {}) };
  for (const [id, m] of Object.entries(b.added || {})) if (!added[id] || (m.at || '') > (added[id].at || '')) added[id] = m;
  const useB = (b.nameAt || '') >= (a.nameAt || '');
  return { person: (useB ? b.person : a.person) || a.person || b.person || '', nameAt: useB ? b.nameAt : a.nameAt, entries: [...entries.values()], added };
}

/** True when `part` has something the file's copy is missing (so the file needs putting right). */
export function partMissingFrom(part, inFile = emptyPart()) {
  const ids = new Set((inFile.entries || []).map((e) => e.id));
  if ((part.entries || []).some((e) => !ids.has(e.id))) return true;
  for (const [id, m] of Object.entries(part.added || {})) {
    const f = inFile.added?.[id];
    if (!f || (m.at || '') > (f.at || '')) return true;
  }
  return (part.nameAt || '') > (inFile.nameAt || '');
}

/** Sections ({ people: { owner: part } }) as a list of per-person logs, the shape the Team log panel uses. */
export function docToLogs(doc) {
  return Object.entries(doc?.people || {}).map(([owner, p]) => ({
    app: APP, owner, person: p.person || 'Unnamed', timesheet: doc.timesheet,
    entries: p.entries || [],
    added: Object.entries(p.added || {}).filter(([, m]) => !m.removed).map(([entryId, m]) => ({ entryId, row: m.row, at: m.at })),
  }));
}

/** Turn an edited log (entries + current markers) back into a section, recording removed markers. */
export function logToPart(log, prev = emptyPart(), name, now = new Date().toISOString()) {
  const added = { ...(prev.added || {}) };
  const current = new Set(log.added.map((m) => m.entryId));
  for (const m of log.added) if (!added[m.entryId] || added[m.entryId].removed) added[m.entryId] = { row: m.row, at: m.at || now };
  for (const id of Object.keys(added)) if (!current.has(id) && !added[id].removed) added[id] = { ...added[id], removed: true, at: now };
  const renamed = name && name !== prev.person;
  return { person: name || prev.person, nameAt: renamed ? now : prev.nameAt || now, entries: log.entries, added };
}
