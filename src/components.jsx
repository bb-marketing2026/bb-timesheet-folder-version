import { useEffect, useRef, useState } from 'react';
import { pad, fmtDate, parseAuDate, localDate } from './core/util.js';

export function NewTimesheet({ onCancel, onCreate }) {
  const [form, setForm] = useState({ client: '', project: '', contact: '', commencement: localDate(new Date()) });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  return (
    <div className="modal" onClick={onCancel}>
      <form className="card dialog" onClick={(e) => e.stopPropagation()}
        onSubmit={(e) => { e.preventDefault(); if (form.client.trim()) onCreate(form); }}>
        <h2>New client timesheet</h2>
        <label className="field"><span>Department / Client</span><input value={form.client} onChange={set('client')} required autoFocus /></label>
        <label className="field"><span>Project name</span><input value={form.project} onChange={set('project')} /></label>
        <label className="field"><span>Point(s) of contact</span><input value={form.contact} onChange={set('contact')} /></label>
        <label className="field"><span>Project commencement</span><DateField value={form.commencement} onChange={(v) => setForm((f) => ({ ...f, commencement: v }))} /></label>
        <div className="row gap end">
          <button type="button" className="ghost" onClick={onCancel}>Cancel</button>
          <button type="submit">Create</button>
        </div>
      </form>
    </div>
  );
}

// Date box that always shows and accepts DD/MM/YYYY. The calendar button uses the
// browser's picker, but the value shown is always Australian format.
export function DateField({ value, onChange }) {
  const [text, setText] = useState(fmtDate(value));
  const [bad, setBad] = useState(false);
  const pickerRef = useRef(null);
  useEffect(() => { setText(fmtDate(value)); setBad(false); }, [value]);
  const commit = () => {
    const key = parseAuDate(text);
    if (key) { setBad(false); if (key !== value) onChange(key); setText(fmtDate(key)); } else setBad(true);
  };
  return (
    <span className={`datefield ${bad ? 'bad' : ''}`}>
      <input type="text" inputMode="numeric" placeholder="DD/MM/YYYY" value={text} aria-label="Date (DD/MM/YYYY)"
        onChange={(e) => setText(e.target.value)} onBlur={commit}
        onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commit(); } }} />
      <button type="button" className="cal" aria-label="Open calendar"
        onClick={() => { try { pickerRef.current?.showPicker(); } catch { pickerRef.current?.focus(); } }}>
        <svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><rect x="1.5" y="3" width="13" height="11.5" rx="2" fill="none" stroke="currentColor" strokeWidth="1.5" /><path d="M1.5 6.5h13M5 1.5v3M11 1.5v3" stroke="currentColor" strokeWidth="1.5" /></svg>
      </button>
      <input ref={pickerRef} type="date" className="picker" tabIndex={-1} aria-hidden="true" value={value}
        onChange={(e) => e.target.value && onChange(e.target.value)} />
    </span>
  );
}

// 24-hour HH:MM box, the same in every browser and region. Accepts 9:5, 0930 or 09:30.
export function TimeField({ value, onChange, label }) {
  const [text, setText] = useState(value);
  const [bad, setBad] = useState(false);
  useEffect(() => { setText(value); setBad(false); }, [value]);
  const commit = () => {
    const m = text.trim().match(/^(\d{1,2})[:.]?(\d{2})$/);
    const h = m ? Number(m[1]) : -1; const mi = m ? Number(m[2]) : -1;
    if (h >= 0 && h < 24 && mi >= 0 && mi < 60) {
      const v = `${pad(h)}:${pad(mi)}`; setBad(false); setText(v); if (v !== value) onChange(v);
    } else setBad(true);
  };
  return (
    <input type="text" inputMode="numeric" className={`timefield ${bad ? 'bad' : ''}`} placeholder="HH:MM" aria-label={`${label} (24 hour)`}
      value={text} onChange={(e) => setText(e.target.value)} onBlur={commit}
      onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commit(); } }} />
  );
}
