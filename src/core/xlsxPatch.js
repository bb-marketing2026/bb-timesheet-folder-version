// Safe, surgical edits to an existing .xlsx.
//
// The file is never rebuilt. Only the cells the app writes are changed in the sheet XML.
// Everything else (other sheets, shared strings, sensitivity labels, SharePoint metadata,
// custom properties, printer settings, calc chain) is left exactly as it was.
// New text is written as inline strings so the shared strings table is never touched.
import JSZip from 'jszip';
import { analyse, readWorkbook, sessionsToRows } from './timesheet.js';

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
const AU_DATE_CODE = 'dd/mm/yyyy';
// Built-in Excel formats that can show month first depending on the computer's region.
const BUILTIN_DATE_FMTS = { 14: 'm/d/yyyy', 22: 'm/d/yyyy h:mm' };

// ---------- small XML helpers ----------
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

function parseAttrs(text) {
  const out = [];
  const re = /([\w:.-]+)="([^"]*)"/g;
  let m;
  while ((m = re.exec(text))) out.push([m[1], m[2]]);
  return out;
}
const attrsToText = (attrs) => attrs.map(([k, v]) => ` ${k}="${v}"`).join('');
const getAttr = (attrs, k) => attrs.find(([a]) => a === k)?.[1];
function setAttr(attrs, k, v) {
  const i = attrs.findIndex(([a]) => a === k);
  if (v === undefined || v === null) { if (i >= 0) attrs.splice(i, 1); return; }
  if (i >= 0) attrs[i][1] = String(v); else attrs.push([k, String(v)]);
}

export function colIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
export function colLetter(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ---------- package ----------
async function openPackage(bytes, sheetName) {
  const zip = await JSZip.loadAsync(bytes);
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const sheetTag = [...wbXml.matchAll(/<sheet\b([^>]*)\/?>/g)]
    .map((m) => parseAttrs(m[1]))
    .find((a) => unesc(getAttr(a, 'name') || '') === sheetName);
  if (!sheetTag) throw new Error(`Sheet "${sheetName}" not found.`);
  const rid = getAttr(sheetTag, 'r:id');
  const rel = [...relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)].map((m) => parseAttrs(m[1])).find((a) => getAttr(a, 'Id') === rid);
  let target = getAttr(rel, 'Target');
  target = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
  const sheetXml = await zip.file(target).async('string');
  const stylesXml = zip.file('xl/styles.xml') ? await zip.file('xl/styles.xml').async('string') : null;
  const date1904 = /<workbookPr\b[^>]*date1904="(1|true)"/.test(wbXml);
  return { zip, wbXml, sheetPath: target, sheetXml, stylesXml, date1904 };
}

// ---------- sheet model: only rows we touch are parsed ----------
class Sheet {
  constructor(xml) {
    const start = xml.indexOf('<sheetData');
    const open = xml.indexOf('>', start) + 1;
    if (xml[open - 2] === '/') { // <sheetData/>
      this.head = `${xml.slice(0, start)}<sheetData>`; this.body = ''; this.tail = `</sheetData>${xml.slice(open)}`;
    } else {
      const close = xml.indexOf('</sheetData>');
      this.head = xml.slice(0, open); this.body = xml.slice(open, close); this.tail = xml.slice(close);
    }
    this.rows = new Map(); // r -> {attrs, cells: Map(col -> {attrs, inner}), raw, start, end}
    this.index = new Map();
    const re = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
    let m;
    while ((m = re.exec(this.body))) {
      const r = Number(getAttr(parseAttrs(m[1]), 'r'));
      this.index.set(r, { start: m.index, end: m.index + m[0].length, head: m[1], inner: m[3] || '' });
    }
    this.newRows = new Set();
  }

  row(r) {
    if (this.rows.has(r)) return this.rows.get(r);
    const ix = this.index.get(r);
    const row = { attrs: ix ? parseAttrs(ix.head) : [['r', String(r)]], cells: new Map(), dirty: false };
    if (ix) {
      const re = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
      let m;
      while ((m = re.exec(ix.inner))) {
        const attrs = parseAttrs(m[1]);
        const col = colIndex(getAttr(attrs, 'r').replace(/\d+/g, ''));
        row.cells.set(col, { attrs, inner: m[3] || '' });
      }
    } else this.newRows.add(r);
    this.rows.set(r, row);
    return row;
  }

  cell(r, c) { return this.row(r).cells.get(c) || null; }

  hasFormula(r, c) { return /<f[\s>/]/.test(this.cell(r, c)?.inner || ''); }

  // Replace a cell's value. kind: 'text' | 'number' | 'blank' | 'raw'
  put(r, c, { kind, value, style, rawType, rawInner }) {
    const row = this.row(r);
    let cell = row.cells.get(c);
    if (!cell) { cell = { attrs: [['r', `${colLetter(c)}${r}`]], inner: '' }; row.cells.set(c, cell); }
    ['t', 'cm', 'vm'].forEach((k) => setAttr(cell.attrs, k, null));
    if (style !== undefined && style !== null) setAttr(cell.attrs, 's', style);
    if (kind === 'text') {
      setAttr(cell.attrs, 't', 'inlineStr');
      cell.inner = `<is><t xml:space="preserve">${esc(value)}</t></is>`;
    } else if (kind === 'number') {
      cell.inner = `<v>${value}</v>`;
    } else if (kind === 'raw') {
      if (rawType) setAttr(cell.attrs, 't', rawType);
      cell.inner = rawInner || '';
    } else cell.inner = '';
    // Keep the r attribute first, as Excel writes it.
    cell.attrs.sort((a, b) => (a[0] === 'r' ? -1 : b[0] === 'r' ? 1 : 0));
    row.dirty = true;
  }

  styleOf(r, c) { return getAttr(this.cell(r, c)?.attrs || [], 's'); }

  // Row height travels with the row's content when rows move.
  heightOf(r) { const a = this.row(r).attrs; return { ht: getAttr(a, 'ht'), customHeight: getAttr(a, 'customHeight') }; }

  setHeight(r, h) {
    const row = this.row(r);
    setAttr(row.attrs, 'ht', h?.ht);
    setAttr(row.attrs, 'customHeight', h?.customHeight);
    row.dirty = true;
  }

  setCachedValue(r, c, value) {
    const cell = this.cell(r, c);
    if (!cell || !/<f[\s>/]/.test(cell.inner)) return;
    cell.inner = /<v>[^<]*<\/v>/.test(cell.inner)
      ? cell.inner.replace(/<v>[^<]*<\/v>/, `<v>${value}</v>`)
      : `${cell.inner}<v>${value}</v>`;
    this.row(r).dirty = true;
  }

  serializeRow(row) {
    const cols = [...row.cells.keys()].sort((a, b) => a - b);
    const spans = getAttr(row.attrs, 'spans');
    if (spans && cols.length) {
      const [lo, hi] = spans.split(':').map(Number);
      if (cols[0] < lo || cols.at(-1) > hi) setAttr(row.attrs, 'spans', null);
    }
    const inner = cols.map((c) => {
      const cell = row.cells.get(c);
      return cell.inner ? `<c${attrsToText(cell.attrs)}>${cell.inner}</c>` : `<c${attrsToText(cell.attrs)}/>`;
    }).join('');
    return inner ? `<row${attrsToText(row.attrs)}>${inner}</row>` : `<row${attrsToText(row.attrs)}/>`;
  }

  toXml() {
    // Rebuild the body: unchanged rows are copied verbatim, new rows slotted in order.
    const existing = [...this.index.entries()].sort((a, b) => a[0] - b[0]);
    const added = [...this.newRows].filter((r) => this.rows.get(r).dirty).sort((a, b) => a - b);
    let out = '';
    let pos = 0;
    let ai = 0;
    for (const [r, ix] of existing) {
      while (ai < added.length && added[ai] < r) { out += this.body.slice(pos, ix.start); pos = ix.start; out += this.serializeRow(this.rows.get(added[ai++])); }
      out += this.body.slice(pos, ix.start);
      const row = this.rows.get(r);
      out += row && row.dirty ? this.serializeRow(row) : this.body.slice(ix.start, ix.end);
      pos = ix.end;
    }
    out += this.body.slice(pos);
    while (ai < added.length) out += this.serializeRow(this.rows.get(added[ai++]));
    return this.head + out + this.tail;
  }
}

// ---------- styles: find or add a DD/MM/YYYY copy of a cell style ----------
class Styles {
  constructor(xml) { this.xml = xml; this.changed = false; this.cache = new Map(); }

  numFmtCode(id) {
    const m = this.xml.match(new RegExp(`<numFmt\\b[^>]*numFmtId="${id}"[^>]*formatCode="([^"]*)"`))
      || this.xml.match(new RegExp(`<numFmt\\b[^>]*formatCode="([^"]*)"[^>]*numFmtId="${id}"`));
    return m ? unesc(m[1]) : BUILTIN_DATE_FMTS[id] || '';
  }

  xfs() {
    const sec = this.xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
    return sec ? [...sec[1].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((m) => m[0]) : [];
  }

  isMonthFirst(styleIndex) {
    const xf = this.xfs()[Number(styleIndex || 0)];
    if (!xf) return false;
    const id = Number(getAttr(parseAttrs(xf.match(/<xf\b([^>]*)/)[1]), 'numFmtId') || 0);
    const code = this.numFmtCode(id).toLowerCase().replace(/\[[^\]]*\]/g, '');
    return /^[^d]*m{1,2}[/.-]d/.test(code);
  }

  auFmtId() {
    const found = this.xml.match(/<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="dd\/mm\/yyyy"/);
    if (found) return Number(found[1]);
    const ids = [...this.xml.matchAll(/numFmtId="(\d+)"/g)].map((m) => Number(m[1]));
    const id = Math.max(163, ...ids) + 1;
    const tag = `<numFmt numFmtId="${id}" formatCode="${AU_DATE_CODE}"/>`;
    if (/<numFmts\b[^>]*>/.test(this.xml)) {
      this.xml = this.xml.replace(/<numFmts\b([^>]*)count="(\d+)"([^>]*)>/, (m, a, n, b) => `<numFmts${a}count="${Number(n) + 1}"${b}>`)
        .replace('</numFmts>', `${tag}</numFmts>`);
    } else {
      this.xml = this.xml.replace(/(<styleSheet\b[^>]*>)/, `$1<numFmts count="1">${tag}</numFmts>`);
    }
    this.changed = true;
    return id;
  }

  // Style index identical to `base` but with DD/MM/YYYY. Reuses an existing one where possible.
  auDateStyle(base) {
    const key = String(base || 0);
    if (this.cache.has(key)) return this.cache.get(key);
    const list = this.xfs();
    const src = list[Number(key)] || list[0];
    const fmtId = this.auFmtId();
    let clone = src.replace(/numFmtId="\d+"/, `numFmtId="${fmtId}"`);
    if (!/applyNumberFormat=/.test(clone)) clone = clone.replace(/<xf\b/, '<xf applyNumberFormat="1"');
    const current = this.xfs();
    let idx = current.indexOf(clone);
    if (idx < 0 && Number(getAttr(parseAttrs(src.match(/<xf\b([^>]*)/)[1]), 'numFmtId')) === fmtId) idx = Number(key);
    if (idx < 0) {
      idx = current.length;
      this.xml = this.xml.replace(/<cellXfs\b([^>]*)count="(\d+)"([^>]*)>/, (m, a, n, b) => `<cellXfs${a}count="${Number(n) + 1}"${b}>`)
        .replace('</cellXfs>', `${clone}</cellXfs>`);
      this.changed = true;
    }
    this.cache.set(key, String(idx));
    return String(idx);
  }
}

function excelSerial(dateKey, date1904) {
  const [y, m, d] = dateKey.split('-').map(Number);
  const serial = (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
  return date1904 ? serial - 1462 : serial;
}

const round2 = (n) => Math.round(n * 100) / 100;

async function savePackage(pkg, sheet, styles) {
  const { zip } = pkg;
  zip.file(pkg.sheetPath, sheet.toXml(), { createFolders: false });
  if (styles?.changed) zip.file('xl/styles.xml', styles.xml, { createFolders: false });
  // Ask Excel to recalculate totals when the file opens.
  if (/<calcPr\b/.test(pkg.wbXml) && !/fullCalcOnLoad=/.test(pkg.wbXml)) {
    zip.file('xl/workbook.xml', pkg.wbXml.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"'), { createFolders: false });
  }
  // Brainbox file standard: Company = Brainbox.
  const app = zip.file('docProps/app.xml');
  if (app) {
    const xml = await app.async('string');
    const next = /<Company>[^<]*<\/Company>|<Company\/>/.test(xml)
      ? xml.replace(/<Company>[^<]*<\/Company>|<Company\/>/, '<Company>Brainbox</Company>')
      : xml.replace('</Properties>', '<Company>Brainbox</Company></Properties>');
    if (next !== xml) zip.file('docProps/app.xml', next, { createFolders: false });
  }
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

function dateColumnToAu(sheet, styles, info) {
  if (!info.cols.date || !styles) return;
  const last = info.totalRow ? info.totalRow - 1 : info.lastUsedRow;
  for (let r = info.firstDataRow; r <= last; r++) {
    const s = sheet.styleOf(r, info.cols.date);
    if (s !== undefined && styles.isMonthFirst(s)) {
      const cell = sheet.cell(r, info.cols.date);
      setAttr(cell.attrs, 's', styles.auDateStyle(s));
      sheet.row(r).dirty = true;
    }
  }
}

/**
 * Add sessions to the timesheet in date order. Returns { bytes, rows }.
 * Usually this is a plain append below the last entry. When a session is dated
 * earlier than the last entry (for example a colleague's work from Monday added
 * on Wednesday), it is slotted in after the last row of its day and the rows
 * below move down. Month and date show on the first row of each day only.
 * Throws a friendly error if the timesheet has too few blank rows.
 */
export async function appendSessionsSafe(bytes, sheetName, sessions, { combine = true } = {}) {
  return placeEntriesSafe(bytes, sheetName, sessionsToRows(sessions, combine));
}

/**
 * Place ready-made rows ({ id?, date, tag, activity, ms }) in date order.
 * Returns { bytes, rows, placed: [{ id, row }] }.
 */
export async function placeEntriesSafe(bytes, sheetName, entries) {
  const info = analyse(await readWorkbook(bytes), sheetName);
  const items = [...entries].sort((a, b) => a.date.localeCompare(b.date));
  if (!items.length) return { bytes, rows: [], placed: [] };
  const K = items.length;
  const left = info.totalRow ? info.totalRow - 1 - info.lastUsedRow : Infinity;
  if (K > left) {
    const err = new Error(`Not enough blank rows. This needs ${K} and the timesheet has ${left} left above "Total Hours". Use "Insert rows" in the file panel, then try again.`);
    err.rowsNeeded = K - left;
    throw err;
  }

  const pkg = await openPackage(bytes, info.sheetName);
  const sheet = new Sheet(pkg.sheetXml);
  const styles = pkg.stylesXml ? new Styles(pkg.stylesXml) : null;
  const { cols } = info;
  const colNums = Object.values(cols);
  const lo = Math.min(...colNums);
  const hi = Math.max(...colNums);

  // Every row in the data area as a unit, with the date it belongs to.
  const byRow = new Map(info.entries.map((e) => [e.row, e]));
  const units = [];
  let carry = '';
  for (let r = info.firstDataRow; r <= info.lastUsedRow; r++) {
    const e = byRow.get(r);
    if (e && e.date) carry = e.date;
    units.push({ row: r, date: carry });
  }
  // Where each new row goes: after the last existing row dated on or before it.
  const posOf = (d) => { let p = 0; units.forEach((u, i) => { if (u.date && u.date <= d) p = i + 1; else if (!u.date) p = Math.max(p, i + 1); }); return p; };
  const placed = items.map((item) => ({ item, pos: posOf(item.date) }));
  const minPos = Math.min(...placed.map((p) => p.pos));

  // Final order of everything from the first affected row down.
  const seq = [];
  for (let i = minPos; i <= units.length; i++) {
    placed.filter((p) => p.pos === i).forEach((p) => seq.push({ kind: 'new', item: p.item, date: p.item.date }));
    if (i < units.length) seq.push({ kind: 'old', unit: units[i], date: units[i].date });
  }
  const startRow = info.firstDataRow + minPos;

  // Existing rows that will move must not hold formulas.
  const moving = units.slice(minPos);
  for (const u of moving) {
    for (let c = lo; c <= hi; c++) {
      if (sheet.hasFormula(u.row, c)) throw new Error(`Row ${u.row} contains a formula, so the app will not move it. Add this entry in Excel instead.`);
    }
  }
  // Snapshot the rows that move before anything is overwritten.
  const snap = new Map(moving.map((u) => [u.row, new Map(
    Array.from({ length: hi - lo + 1 }, (_, k) => lo + k).map((c) => {
      const cell = sheet.cell(u.row, c);
      return [c, cell ? { t: getAttr(cell.attrs, 't'), s: getAttr(cell.attrs, 's'), inner: cell.inner } : null];
    }),
  )]));
  const heights = new Map(moving.map((u) => [u.row, sheet.heightOf(u.row)]));
  const styleRow = Math.max(info.firstDataRow, info.lastUsedRow);

  let prevDate = minPos > 0 ? units[minPos - 1].date : '';
  const rows = [];
  const placedRows = [];
  let added = 0;
  seq.forEach((el, i) => {
    const r = startRow + i;
    if (el.kind === 'old') {
      const cells = snap.get(el.unit.row);
      if (el.unit.row !== r) sheet.setHeight(r, heights.get(el.unit.row));
      for (let c = lo; c <= hi; c++) {
        const src = cells.get(c);
        if (el.unit.row === r) continue; // not moving
        sheet.put(r, c, src && src.inner
          ? { kind: 'raw', rawType: src.t, rawInner: src.inner, style: src.s }
          : { kind: 'blank', style: src ? src.s : sheet.styleOf(r, c) });
      }
    } else {
      const { item } = el;
      if (r <= info.lastUsedRow) sheet.setHeight(r, null); // a moved-down row's height should not stick to the new entry
      const hours = round2(item.ms / 3600000);
      const firstOfDay = item.date !== prevDate;
      // Styles come from the row's own cells when it is a template blank row, otherwise from a normal data row.
      const styleFor = (c) => (r > info.lastUsedRow ? sheet.styleOf(r, c) : undefined) ?? snap.get(styleRow)?.get(c)?.s ?? sheet.styleOf(styleRow, c);
      if (cols.month) sheet.put(r, cols.month, firstOfDay
        ? { kind: 'text', value: MONTHS[Number(item.date.slice(5, 7)) - 1], style: styleFor(cols.month) }
        : { kind: 'blank', style: styleFor(cols.month) });
      if (cols.date) {
        const base = styleFor(cols.date);
        const style = styles ? styles.auDateStyle(base) : base;
        sheet.put(r, cols.date, firstOfDay
          ? { kind: 'number', value: excelSerial(item.date, pkg.date1904), style }
          : { kind: 'blank', style });
      }
      if (cols.tag) sheet.put(r, cols.tag, { kind: 'text', value: item.tag, style: styleFor(cols.tag) });
      if (cols.activity) sheet.put(r, cols.activity, { kind: 'text', value: item.activity, style: styleFor(cols.activity) });
      sheet.put(r, cols.hours, { kind: 'number', value: hours, style: styleFor(cols.hours) });
      if (cols.eod) sheet.put(r, cols.eod, { kind: 'blank', style: styleFor(cols.eod) });
      added += hours;
      rows.push(r);
      placedRows.push({ id: item.id, row: r });
    }
    prevDate = el.date;
  });

  dateColumnToAu(sheet, styles, info);
  if (info.totalRow) sheet.setCachedValue(info.totalRow, cols.hours, round2(info.total + added));
  const out = await savePackage(pkg, sheet, styles);

  // Check before anything is saved: every existing entry is still there, in order, and totals add up.
  const after = analyse(await readWorkbook(out), info.sheetName);
  const k = (e) => `${e.tag}|${e.activity}|${round2(e.hours)}`;
  const newRows = new Set(rows);
  const kept = after.entries.filter((e) => !newRows.has(e.row)).map(k);
  const problems = [];
  if (kept.join('\n') !== info.entries.map(k).join('\n')) problems.push('existing entries changed');
  if (round2(after.total) !== round2(info.total + added)) problems.push('hours total does not add up');
  if (after.entries.length !== info.entries.length + K) problems.push('row count is wrong');
  const dates = after.entries.map((e) => e.date).filter(Boolean);
  const orderedBefore = info.entries.every((e, i, a) => !i || !e.date || !a[i - 1].date || a[i - 1].date <= e.date);
  if (orderedBefore && dates.some((d, i) => i && dates[i - 1] > d)) problems.push('dates out of order');
  if (problems.length) throw new Error(`Adding was stopped and nothing was saved (${problems.join('; ')}).`);
  return { bytes: out, rows, placed: placedRows };
}

/**
 * Delete one entry. Rows below move up one, so there are no gaps, and the
 * Total Hours formula range is unchanged. `expected` must match what is in the
 * file right now, otherwise nothing is changed.
 */
export async function deleteEntrySafe(bytes, sheetName, expected) {
  const info = analyse(await readWorkbook(bytes), sheetName);
  const current = info.entries.find((e) => e.row === expected.row);
  const same = current && current.tag === expected.tag && current.activity === expected.activity
    && round2(current.hours) === round2(expected.hours);
  if (!same) throw new Error('That row has changed in the file since it was shown here, so nothing was deleted. The list has been refreshed.');

  const pkg = await openPackage(bytes, info.sheetName);
  const sheet = new Sheet(pkg.sheetXml);
  const styles = pkg.stylesXml ? new Styles(pkg.stylesXml) : null;
  const colNums = Object.values(info.cols);
  const lo = Math.min(...colNums);
  const hi = Math.max(...colNums);
  const R = expected.row;
  const last = info.lastUsedRow;

  for (let r = R; r <= last; r++) {
    for (let c = lo; c <= hi; c++) {
      if (sheet.hasFormula(r, c)) throw new Error(`Row ${r} contains a formula, so the app will not move it. Delete this row in Excel instead.`);
    }
  }

  // If the deleted row carries the day's month and date, hand them to the next row of the same day.
  const { month: mc, date: dc } = info.cols;
  const next = info.entries.find((e) => e.row > R);
  if (dc && current.hasDate && next && next.date === current.date && !next.hasDate) {
    const copy = (c) => {
      const src = sheet.cell(R, c);
      sheet.put(next.row, c, { kind: 'raw', rawType: getAttr(src?.attrs || [], 't'), rawInner: src?.inner || '', style: sheet.styleOf(next.row, c) });
    };
    if (mc) copy(mc);
    copy(dc);
  }

  // Move every cell in the data block up one row, then clear the old last row.
  for (let r = R; r < last; r++) sheet.setHeight(r, sheet.heightOf(r + 1));
  sheet.setHeight(last, null);
  for (let r = R; r < last; r++) {
    for (let c = lo; c <= hi; c++) {
      const src = sheet.cell(r + 1, c);
      sheet.put(r, c, src && src.inner
        ? { kind: 'raw', rawType: getAttr(src.attrs, 't'), rawInner: src.inner, style: getAttr(src.attrs, 's') }
        : { kind: 'blank', style: src ? getAttr(src.attrs, 's') : sheet.styleOf(r, c) });
    }
  }
  for (let c = lo; c <= hi; c++) sheet.put(last, c, { kind: 'blank', style: sheet.styleOf(last, c) });

  dateColumnToAu(sheet, styles, info);
  if (info.totalRow) sheet.setCachedValue(info.totalRow, info.cols.hours, round2(info.total - current.hours));
  return { bytes: await savePackage(pkg, sheet, styles), removed: current };
}

// ---------- inserting rows ----------
// Shifts cell references in a formula when N rows are inserted at row P of `target`.
// Ranges that end on the last row before the Total row (P - 1) and start inside the
// data area are extended, so SUM(F11:F146) becomes SUM(F11:F196).
const REF_RE = /("(?:[^"]|"")*")|(?<![\w.$])((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(?:(\$?[A-Z]{1,3}\$?)(\d+)(?::(\$?[A-Z]{1,3}\$?)(\d+))?|(\$?)(\d+):(\$?)(\d+))(?![\w(!])/g;

export function shiftFormula(f, { P, N, target, here, extendFrom }) {
  return f.replace(REF_RE, (m, str, prefix, c1, r1, c2, r2, d1, rr1, d2, rr2) => {
    if (str) return m;
    const sheet = prefix ? prefix.slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'") : (here ? target : null);
    if (sheet !== target) return m;
    const pre = prefix || '';
    const shift = (r, isEnd, start) => {
      let n = Number(r);
      if (isEnd && n === P - 1 && start !== undefined && start >= extendFrom && start <= P - 1) return n + N;
      if (n >= P) n += N;
      return n;
    };
    if (c1 !== undefined) {
      if (c2 === undefined) return `${pre}${c1}${shift(r1)}`;
      return `${pre}${c1}${shift(r1)}:${c2}${shift(r2, true, Number(r1))}`;
    }
    return `${pre}${d1}${shift(rr1)}:${d2}${shift(rr2, true, Number(rr1))}`;
  });
}

const shiftSqref = (sq, ctx) => sq.split(/\s+/).map((t) => shiftFormula(t, ctx)).join(' ');

function shiftAttrs(xml, tag, attr, fn) {
  return xml.replace(new RegExp(`(<${tag}\\b[^>]*?\\b${attr}=")([^"]*)(")`, 'g'), (m, a, v, b) => `${a}${fn(v)}${b}`);
}

function shiftRowXml(rowXml, P, N, ctx) {
  // Renumber the row and its cells, and shift formula text and shared/array formula ranges.
  return rowXml
    .replace(/^<row\b([^>]*?)\br="(\d+)"/, (m, a, r) => `<row${a}r="${Number(r) >= P ? Number(r) + N : r}"`)
    .replace(/<c\b([^>]*?)\br="([A-Z]{1,3})(\d+)"/g, (m, a, c, r) => `<c${a}r="${c}${Number(r) >= P ? Number(r) + N : r}"`)
    .replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/g, (m, a, body) => `<f${a.replace(/\bref="([^"]*)"/, (x, v) => `ref="${shiftSqref(v, ctx)}"`)}>${esc(shiftFormula(unesc(body), ctx))}</f>`)
    .replace(/<f\b([^>]*?)\/>/g, (m, a) => `<f${a.replace(/\bref="([^"]*)"/, (x, v) => `ref="${shiftSqref(v, ctx)}"`)}/>`);
}

/**
 * Insert `count` blank, formatted rows directly above the Total Hours row.
 * Refuses (changes nothing) if the workbook has anything the app cannot shift safely.
 */
export async function insertRowsSafe(bytes, sheetName, count) {
  const N = Math.floor(Number(count));
  if (!Number.isFinite(N) || N < 1 || N > 500) throw new Error('Enter a number of rows from 1 to 500.');
  const info = analyse(await readWorkbook(bytes), sheetName);
  if (!info.totalRow) throw new Error('No "Total Hours" row was found, so the app does not know where to insert rows.');
  const P = info.totalRow;
  const pkg = await openPackage(bytes, info.sheetName);
  const { zip } = pkg;
  const target = info.sheetName;
  const ctx = { P, N, target, here: true, extendFrom: info.firstDataRow };
  const other = { ...ctx, here: false };

  // Features the app will not touch. Excel handles these properly, so send the user there.
  const relsPath = pkg.sheetPath.replace(/worksheets\/([^/]+)$/, 'worksheets/_rels/$1.rels');
  const rels = zip.file(relsPath) ? await zip.file(relsPath).async('string') : '';
  const blockers = [];
  if (/<tableParts\b/.test(pkg.sheetXml) || /relationships\/table"/.test(rels)) blockers.push('a formatted table');
  if (/<drawing\b|<legacyDrawing\b/.test(pkg.sheetXml) || /relationships\/(drawing|vmlDrawing|comments)"/.test(rels)) blockers.push('charts, images or comments');
  if (/relationships\/pivotTable"/.test(rels) || Object.keys(zip.files).some((n) => n.startsWith('xl/pivotCache'))) blockers.push('a pivot table');
  if (/<extLst>[\s\S]*xm:sqref/.test(pkg.sheetXml)) blockers.push('advanced conditional formatting or validation');
  if (/<f\b[^>]*\bt="(shared|array)"[^>]*\bref="([^"]*)"/.test(pkg.sheetXml)) {
    for (const m of pkg.sheetXml.matchAll(/<f\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g)) {
      if (Number(m[2]) < P && Number(m[4]) >= P) { blockers.push('a shared formula across the Total row'); break; }
    }
  }
  for (const name of Object.keys(zip.files)) {
    if (/^xl\/charts\/.*\.xml$/.test(name) && (await zip.file(name).async('string')).includes(`${target}'!`)) { blockers.push('a chart using this sheet'); break; }
  }
  if (blockers.length) throw new Error(`This sheet has ${blockers.join(', ')}, so the app will not insert rows. Insert them in Excel instead: right-click row ${P}, Insert.`);

  // 1. Sheet rows.
  const sd0 = pkg.sheetXml.indexOf('<sheetData');
  const sdOpen = pkg.sheetXml.indexOf('>', sd0) + 1;
  const sdClose = pkg.sheetXml.indexOf('</sheetData>');
  let head = pkg.sheetXml.slice(0, sdOpen);
  const body = pkg.sheetXml.slice(sdOpen, sdClose);
  let tail = pkg.sheetXml.slice(sdClose);
  const rowRe = /<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g;
  let out = '';
  let pos = 0;
  let inserted = false;
  let template = null;
  const blankRows = () => {
    const t = template;
    return Array.from({ length: N }, (_, i) => {
      const r = P + i;
      if (!t) return `<row r="${r}"/>`;
      const attrs = parseAttrs(t.match(/^<row\b([^>]*?)\/?>/)[1]);
      setAttr(attrs, 'r', r);
      const cells = [...t.matchAll(/<c\b([^>]*?)(?:\/>|>[\s\S]*?<\/c>)/g)].map((m) => {
        const a = parseAttrs(m[1]).filter(([k]) => k === 'r' || k === 's');
        setAttr(a, 'r', `${getAttr(a, 'r').replace(/\d+/g, '')}${r}`);
        return `<c${attrsToText(a)}/>`;
      }).join('');
      return cells ? `<row${attrsToText(attrs)}>${cells}</row>` : `<row${attrsToText(attrs)}/>`;
    }).join('');
  };
  for (const m of body.matchAll(rowRe)) {
    const r = Number(m[0].match(/^<row\b[^>]*?\br="(\d+)"/)[1]);
    out += body.slice(pos, m.index);
    if (r === P - 1) template = m[0];
    if (r >= P && !inserted) { out += blankRows(); inserted = true; }
    out += r >= P || /<f\b/.test(m[0]) ? shiftRowXml(m[0], P, N, ctx) : m[0];
    pos = m.index + m[0].length;
  }
  out += body.slice(pos);
  if (!inserted) out += blankRows();

  // 2. Everything outside the rows that refers to row numbers.
  const fixOutside = (xml) => {
    let x = xml;
    x = shiftAttrs(x, 'dimension', 'ref', (v) => shiftSqref(v, ctx));
    x = shiftAttrs(x, 'autoFilter', 'ref', (v) => shiftSqref(v, ctx));
    x = shiftAttrs(x, 'mergeCell', 'ref', (v) => shiftSqref(v, ctx));
    x = shiftAttrs(x, 'hyperlink', 'ref', (v) => shiftSqref(v, ctx));
    x = shiftAttrs(x, 'conditionalFormatting', 'sqref', (v) => shiftSqref(v, ctx));
    x = shiftAttrs(x, 'dataValidation', 'sqref', (v) => shiftSqref(v, ctx));
    x = shiftAttrs(x, 'selection', 'sqref', (v) => shiftSqref(v, ctx));
    x = shiftAttrs(x, 'selection', 'activeCell', (v) => shiftSqref(v, ctx));
    x = shiftAttrs(x, 'pane', 'topLeftCell', (v) => shiftSqref(v, ctx));
    x = x.replace(/<(formula|formula1|formula2)>([\s\S]*?)<\/\1>/g, (m, t, f) => `<${t}>${esc(shiftFormula(unesc(f), ctx))}</${t}>`);
    x = x.replace(/(<rowBreaks\b[\s\S]*?<\/rowBreaks>)/, (blk) => blk.replace(/<brk\b([^>]*?)\bid="(\d+)"/g, (m, a, id) => `<brk${a}id="${Number(id) >= P - 1 ? Number(id) + N : id}"`));
    return x;
  };
  head = fixOutside(head);
  tail = fixOutside(tail);
  zip.file(pkg.sheetPath, head + out + tail, { createFolders: false });

  // 3. Defined names (print area, filter range) and formulas on other sheets that point here.
  let wbXml = pkg.wbXml.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (m, a, f, b) => `${a}${esc(shiftFormula(unesc(f), other))}${b}`);
  if (/<calcPr\b/.test(wbXml) && !/fullCalcOnLoad=/.test(wbXml)) wbXml = wbXml.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
  zip.file('xl/workbook.xml', wbXml, { createFolders: false });
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(name) || name === pkg.sheetPath) continue;
    const xml = await zip.file(name).async('string');
    if (!xml.includes('<f') || !xml.includes(`${target}`)) continue;
    const next = xml.replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/g, (m, a, f) => `<f${a}>${esc(shiftFormula(unesc(f), other))}</f>`);
    if (next !== xml) zip.file(name, next, { createFolders: false });
  }

  // 4. Excel's calculation order list.
  const cc = zip.file('xl/calcChain.xml');
  if (cc) {
    const sheetId = getAttr([...pkg.wbXml.matchAll(/<sheet\b([^>]*)\/?>/g)].map((m) => parseAttrs(m[1])).find((a) => unesc(getAttr(a, 'name') || '') === target), 'sheetId');
    let current = null;
    const xml = await cc.async('string');
    const next = xml.replace(/<c\b([^>]*?)\/>/g, (m, a) => {
      const attrs = parseAttrs(a);
      if (getAttr(attrs, 'i')) current = getAttr(attrs, 'i');
      if (current === sheetId) {
        const ref = getAttr(attrs, 'r');
        const mm = ref.match(/^([A-Z]+)(\d+)$/);
        if (mm && Number(mm[2]) >= P) setAttr(attrs, 'r', `${mm[1]}${Number(mm[2]) + N}`);
      }
      return `<c${attrsToText(attrs)}/>`;
    });
    if (next !== xml) zip.file('xl/calcChain.xml', next, { createFolders: false });
  }

  // Brainbox file standard.
  const app = zip.file('docProps/app.xml');
  if (app) {
    const xml = await app.async('string');
    const next = /<Company>[^<]*<\/Company>|<Company\/>/.test(xml)
      ? xml.replace(/<Company>[^<]*<\/Company>|<Company\/>/, '<Company>Brainbox</Company>')
      : xml.replace('</Properties>', '<Company>Brainbox</Company></Properties>');
    if (next !== xml) zip.file('docProps/app.xml', next, { createFolders: false });
  }
  const result = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });

  // 5. Check the result before anything is saved.
  const after = analyse(await readWorkbook(result), target);
  const key = (e) => `${e.date}|${e.tag}|${e.activity}|${round2(e.hours)}|${e.row}`;
  const problems = [];
  if (after.totalRow !== P + N) problems.push('Total Hours row did not move as expected');
  if (after.rowsLeft !== info.rowsLeft + N) problems.push('blank row count is wrong');
  if (after.entries.map(key).join('\n') !== info.entries.map(key).join('\n')) problems.push('existing entries changed');
  if (round2(after.total) !== round2(info.total)) problems.push('hours total changed');
  if (problems.length) throw new Error(`Inserting rows was stopped and nothing was saved (${problems.join('; ')}). Insert rows in Excel instead.`);
  return { bytes: result, totalRow: after.totalRow, rowsLeft: after.rowsLeft };
}

/** Short SHA-256 of the file, used to spot changes made outside the app. */
export async function fingerprint(bytes) {
  const buf = bytes instanceof ArrayBuffer ? bytes : bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const hash = await crypto.subtle.digest('SHA-256', buf);
  return [...new Uint8Array(hash)].slice(0, 12).map((b) => b.toString(16).padStart(2, '0')).join('');
}

// For one-off maintenance scripts (see scripts/). Not used by the app itself.
export const internal = { openPackage, savePackage, Sheet, Styles, excelSerial };
