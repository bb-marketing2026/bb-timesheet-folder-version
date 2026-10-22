// Shared helpers: Australian dates in Perth time, file names, downloads.
export const pad = (n) => String(n).padStart(2, '0');
export const TZ = 'Australia/Perth';
const PERTH_OFFSET = '+08:00'; // Perth has no daylight saving

const perthParts = (d) => Object.fromEntries(
  new Intl.DateTimeFormat('en-AU', {
    timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23',
  }).formatToParts(d).map((p) => [p.type, p.value]),
);
export const localDate = (d) => { const p = perthParts(d); return `${p.year}-${p.month}-${p.day}`; };
export const localTime = (iso) => { const p = perthParts(new Date(iso)); return `${p.hour}:${p.minute}`; };
export const perthIso = (dateKey, hhmm) => new Date(`${dateKey}T${hhmm}:00${PERTH_OFFSET}`).toISOString();
export const fmtDate = (key) => (key ? key.split('-').reverse().join('/') : '');
export const fmtStamp = (ms) => `${fmtDate(localDate(new Date(ms)))} ${localTime(new Date(ms).toISOString())}`;
export const longDate = (d) => `${new Intl.DateTimeFormat('en-AU', { timeZone: TZ, weekday: 'long' }).format(d)} ${fmtDate(localDate(d))}`;
export const fmtHours = (h) => (Math.round(h * 100) / 100).toFixed(2);
export const uid = () => (crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`);

export const parseAuDate = (text) => {
  const m = String(text).trim().match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})$/);
  if (!m) return '';
  const y = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
  const mo = Number(m[2]); const d = Number(m[1]);
  const dt = new Date(Date.UTC(y, mo - 1, d));
  if (dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return '';
  return `${y}-${pad(mo)}-${pad(d)}`;
};

export const clock = (ms) => {
  const s = Math.max(0, Math.floor(ms / 1000));
  return `${pad(Math.floor(s / 3600))}:${pad(Math.floor((s % 3600) / 60))}:${pad(s % 60)}`;
};

export const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export const XLSX_TYPE = { description: 'Excel timesheet', accept: { [XLSX_MIME]: ['.xlsx'] } };

export function downloadBytes(buffer, filename, type = XLSX_MIME) {
  const url = URL.createObjectURL(new Blob([buffer], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// Brainbox file naming: [Description] - vYYYYMMDD-rHHMM, Perth time.
export function versionedName(name, label = '') {
  const base = name.replace(/\.xlsx$/i, '').replace(/\s*-\s*v\d{8}(-r\d{4})?$/i, '').replace(/\s*-\s*Backup$/i, '')
    + (label ? ` - ${label}` : '');
  const p = perthParts(new Date());
  return `${base} - v${p.year}${p.month}${p.day}-r${p.hour}${p.minute}.xlsx`;
}

// Save bytes where the person chooses (Chrome/Edge), otherwise download.
export async function saveBytesAs(bytes, suggestedName, startIn) {
  if ('showSaveFilePicker' in window) {
    const opts = { suggestedName, types: [XLSX_TYPE] };
    if (startIn) opts.startIn = startIn;
    let handle;
    try { handle = await window.showSaveFilePicker(opts); } catch (e) {
      if (e.name === 'AbortError' || !startIn) throw e;
      handle = await window.showSaveFilePicker({ suggestedName, types: [XLSX_TYPE] });
    }
    const w = await handle.createWritable();
    await w.write(bytes);
    await w.close();
    return { name: handle.name, handle };
  }
  downloadBytes(bytes, suggestedName);
  return { name: suggestedName, handle: null };
}

export class StoreError extends Error {
  constructor(code, message) { super(message); this.code = code; }
}
