// Excel logic for Brainbox timesheets. No React here, so it can be tested in Node.
import ExcelJS from 'exceljs';

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];

// Header text, matched loosely so the different Brainbox timesheet layouts all work
// (for example RAWCS: Month, Date Log, Project Description, Activities, Total Hours, EOD Log;
// Weld East: Month, Date Log, Day, Project Description, Activities, Total Hours, Notes).
// Columns can be in any order; only Total Hours plus Activities or Project Description are required.
const HEADER_MATCHERS = {
  month: /^month$/i,
  day: /^(day|weekday|day of week)$/i,
  date: /^date/i,
  tag: /project description|activity tag|^tag$|category|^(project|service|task type|work type)$/i,
  activity: /^activit|^tasks?$|^task description|description of work|^(description|details|work done)$/i,
  hours: /total hours|^(hours|hrs|time spent|duration)$/i,
  eod: /eod|^(notes?|comments?|remarks?)$/i,
};
// Columns whose content decides whether a row is used. Month, Date and Day are left out because
// some templates fill them in ahead of time (Weld East has weekday names in its blank rows).
const CONTENT_COLS = ['tag', 'activity', 'hours', 'eod'];

export const AU_DATE = 'dd/mm/yyyy';

export function cellText(cell) {
  const v = cell?.value;
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if (v.text !== undefined) return String(v.text);
    if (v.result !== undefined) return String(v.result);
    if (v.formula) return '';
    return '';
  }
  return String(v);
}

const isBlank = (cell) => cellText(cell).trim() === '';

function cellNumber(cell) {
  const v = cell?.value;
  if (typeof v === 'number') return v;
  if (v && typeof v === 'object' && typeof v.result === 'number') return v.result;
  const n = parseFloat(cellText(cell));
  return Number.isFinite(n) ? n : 0;
}

// Dates in the file are a mix of real dates and text such as " 08/10/2026" (dd/mm/yyyy).
function cellDateKey(cell) {
  const v = cell?.value;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const t = cellText(cell).trim();
  const m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const y = m[3].length === 2 ? `20${m[3]}` : m[3];
    return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  }
  return '';
}

function colLetter(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function findHeader(ws) {
  const max = Math.min(ws.rowCount, 40);
  for (let r = 1; r <= max; r++) {
    const row = ws.getRow(r);
    const cols = {};
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      const t = cellText(cell).trim();
      for (const [key, re] of Object.entries(HEADER_MATCHERS)) {
        if (!cols[key] && re.test(t)) { cols[key] = c; break; }
      }
    });
    if (cols.hours && (cols.activity || cols.tag)) return { headerRow: r, cols };
  }
  return null;
}

function findTotalRow(ws, headerRow, hoursCol) {
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    let found = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (/^\s*total(\s+hours|\s+hrs)?\s*:?\s*$/i.test(cellText(cell))) found = true;
    });
    if (found) return r;
  }
  // No label: the first row below the header whose hours cell adds up the column.
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const v = ws.getRow(r).getCell(hoursCol).value;
    if (v && typeof v === 'object' && /^\s*(SUM|SUBTOTAL)\(/i.test(v.formula || v.sharedFormula || '')) return r;
  }
  return null;
}

/** Sheets in the workbook that look like a timesheet. */
export function timesheetSheets(wb) {
  return wb.worksheets.filter((ws) => findHeader(ws)).map((ws) => ws.name);
}

function pickSheet(wb, preferred) {
  const names = timesheetSheets(wb);
  if (!names.length) return null;
  if (preferred && names.includes(preferred)) return preferred;
  const year = String(new Date().getFullYear());
  if (names.includes(year)) return year;
  return names[names.length - 1];
}

function labelValue(ws, label) {
  for (let r = 1; r <= 9; r++) {
    const row = ws.getRow(r);
    let out = '';
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      if (!out && label.test(cellText(cell).trim())) out = cellText(row.getCell(c + 1)).trim();
    });
    if (out) return out;
  }
  return '';
}

/** Read everything the app needs to show from one sheet. */
export function analyse(wb, preferredSheet) {
  const sheetName = pickSheet(wb, preferredSheet);
  if (!sheetName) {
    throw new Error('No timesheet found in this file. It needs a header row with "Project Description", "Activities" and "Total Hours".');
  }
  const ws = wb.getWorksheet(sheetName);
  const { headerRow, cols } = findHeader(ws);
  const firstDataRow = headerRow + 1;
  const totalRow = findTotalRow(ws, headerRow, cols.hours);
  const lastDataRow = totalRow ? totalRow - 1 : ws.rowCount;
  const dataCols = CONTENT_COLS.map((k) => cols[k]).filter(Boolean);

  let lastUsedRow = headerRow;
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    const row = ws.getRow(r);
    if (dataCols.some((c) => !isBlank(row.getCell(c)))) lastUsedRow = r;
  }

  const tagCount = new Map();
  const entries = [];
  let currentDate = '';
  let total = 0;
  for (let r = firstDataRow; r <= lastUsedRow; r++) {
    const row = ws.getRow(r);
    const ownDate = cols.date ? cellDateKey(row.getCell(cols.date)) : '';
    if (ownDate) currentDate = ownDate;
    const tag = cols.tag ? cellText(row.getCell(cols.tag)).trim() : '';
    const activity = cols.activity ? cellText(row.getCell(cols.activity)).trim() : '';
    const hours = cellNumber(row.getCell(cols.hours));
    if (!tag && !activity && !hours) continue;
    if (tag) tagCount.set(tag, (tagCount.get(tag) || 0) + 1);
    total += hours;
    entries.push({ row: r, date: currentDate, hasDate: !!ownDate, tag, activity, hours });
  }

  return {
    sheetName,
    sheets: timesheetSheets(wb),
    client: labelValue(ws, /client/i),
    project: labelValue(ws, /project name/i),
    contact: labelValue(ws, /contact/i),
    headerRow, cols, firstDataRow, totalRow, lastUsedRow,
    lastDate: currentDate,
    tags: [...tagCount.entries()].sort((a, b) => b[1] - a[1]).map(([t]) => t),
    entries,
    total: Math.round(total * 100) / 100,
    rowsLeft: totalRow ? Math.max(0, totalRow - 1 - lastUsedRow) : null,
  };
}

/** Hours between two ISO times, rounded to two decimals like the existing file. */
export function hoursBetween(startIso, endIso) {
  const ms = new Date(endIso) - new Date(startIso);
  return Math.max(0, Math.round((ms / 3600000) * 100) / 100);
}

/** Turn sessions into timesheet rows. Optionally combines sessions with the same tag on the same day. */
export function sessionsToRows(sessions, combine) {
  const sorted = [...sessions].sort((a, b) => new Date(a.start) - new Date(b.start));
  if (!combine) {
    return sorted.map((s) => ({ date: s.date, tag: s.tag.trim(), activity: s.activity.trim(), ms: new Date(s.end) - new Date(s.start) }));
  }
  const map = new Map();
  for (const s of sorted) {
    const key = `${s.date}|${s.tag.trim().toLowerCase()}`;
    const ms = new Date(s.end) - new Date(s.start);
    if (!map.has(key)) map.set(key, { date: s.date, tag: s.tag.trim(), activities: [], ms: 0 });
    const g = map.get(key);
    g.ms += ms;
    const a = s.activity.trim();
    if (a && !g.activities.some((x) => x.toLowerCase() === a.toLowerCase())) g.activities.push(a);
  }
  return [...map.values()].map((g) => ({ date: g.date, tag: g.tag, activity: g.activities.join('; '), ms: g.ms }));
}

export async function readWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

export async function writeWorkbook(wb) {
  wb.company = 'Brainbox';
  wb.modified = new Date();
  return wb.xlsx.writeBuffer();
}

/** A blank timesheet in the Brainbox layout, for a new client. */
export async function createTimesheet({ client, project, contact, commencement }) {
  const PINK = 'FFFF0080';
  const wb = new ExcelJS.Workbook();
  wb.company = 'Brainbox';
  wb.creator = 'Brainbox';
  wb.title = '';
  const ws = wb.addWorksheet(String(new Date().getFullYear()), { views: [{ showGridLines: false, state: 'frozen', ySplit: 10 }] });
  const widths = { A: 2, B: 9, C: 19, D: 36, E: 46, F: 11, G: 32 };
  Object.entries(widths).forEach(([k, w]) => { ws.getColumn(k).width = w; });
  const base = { name: 'Aptos Light', size: 9 };
  const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };
  const fillPink = { type: 'pattern', pattern: 'solid', fgColor: { argb: PINK } };

  ws.getRow(1).height = 34.5;
  const title = ws.getCell('C1');
  title.value = 'Brainbox Project Timesheet';
  title.font = { name: 'Exo SemiBold', size: 22, color: { argb: PINK } };
  title.alignment = { vertical: 'middle' };

  const info = [
    ['Department / Client', client],
    ['Project Name', project],
    ['Project Commencement', commencement ? new Date(`${commencement}T00:00:00Z`) : ''],
    ['Point(s) of Contact', contact],
    ['Project End', ''],
  ];
  info.forEach(([label, value], i) => {
    const r = 3 + i;
    const l = ws.getCell(`B${r}`);
    l.value = label; l.font = { ...base, color: { argb: 'FFFFFFFF' } }; l.fill = fillPink; l.border = box;
    l.alignment = { vertical: 'middle' };
    const v = ws.getCell(`C${r}`);
    v.value = value || null; v.font = { ...base, bold: true, color: { argb: PINK } };
    v.alignment = { horizontal: 'left', vertical: 'middle' }; v.border = { bottom: thin };
    if (value instanceof Date) v.numFmt = 'dd/mm/yyyy';
  });
  ws.getColumn('B').width = 22;

  const headers = ['Month', 'Date Log', 'Project Description', 'Activities', 'Total Hours', 'EOD Log'];
  headers.forEach((h, i) => {
    const c = ws.getRow(10).getCell(2 + i);
    c.value = h; c.font = { ...base, bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = fillPink; c.border = box;
    c.alignment = { horizontal: i === 0 || i === 1 || i === 4 ? 'center' : 'left', vertical: 'middle' };
  });

  const dataRows = 100;
  for (let r = 11; r < 11 + dataRows; r++) {
    for (let c = 2; c <= 7; c++) {
      const cell = ws.getRow(r).getCell(c);
      cell.font = { ...base, bold: c >= 4 };
      cell.border = box;
      cell.alignment = { horizontal: c === 2 ? 'center' : 'left', vertical: 'middle', wrapText: c === 5 };
      if (c === 3) cell.numFmt = 'dd/mm/yyyy';
      if (c === 6) cell.numFmt = '0.00';
    }
  }
  const totalRow = 11 + dataRows;
  const tl = ws.getRow(totalRow).getCell(5);
  tl.value = 'Total Hours'; tl.font = { ...base, bold: true, color: { argb: 'FFFFFFFF' } }; tl.fill = fillPink;
  tl.border = { ...box, bottom: { style: 'medium', color: { argb: PINK } } };
  const tv = ws.getRow(totalRow).getCell(6);
  tv.value = { formula: `SUM(F11:F${totalRow - 1})`, result: 0 };
  tv.font = { ...base, bold: true }; tv.numFmt = '0.00'; tv.alignment = { horizontal: 'center', vertical: 'middle' };
  tv.border = { ...box, bottom: { style: 'medium', color: { argb: PINK } } };
  return wb;
}
