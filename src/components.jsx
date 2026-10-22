import { useEffect, useRef, useState } from 'react';
import { pad, fmtDate, parseAuDate, localDate } from './core/util.js';
import { manual } from './manual.js';

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

// How to use the tracker: opened with the "How to use" button beside User manual.
export function Guide({ edition, onManual, onClose }) {
  const m365 = edition === 'm365';
  return (
    <section className="guide" aria-label="How to use the tracker">
      <div className="guide-head">
        <h2>How to use the tracker</h2>
        <button className="close" onClick={onClose} aria-label="Close how to use">×</button>
      </div>
      <ol>
        <li><strong>Your name and client tag.</strong> Type your name and click Save (once). Colleagues see it in the Team log, never in the timesheet. Beside it, check the Client tag for the open timesheet, for example RAWCS.</li>
        {m365 ? (
          <li><strong>Open the timesheet.</strong> Click Sign in with Microsoft. In SharePoint or OneDrive, click Copy link on the client timesheet, paste it here and click Open. Next time, use Recent clients.</li>
        ) : (
          <li><strong>Open the timesheet.</strong> Click Open timesheet and choose the client's .xlsx file. Next time, use Recent clients.</li>
        )}
        {m365 ? (
          <li><strong>Team log.</strong> Connected automatically. Everyone who opens the same timesheet shares it.</li>
        ) : (
          <li><strong>Team log (once per timesheet).</strong> In the Team log panel, click Connect team log. The folder picker opens where the timesheet is saved. Click Select folder, then Edit files.</li>
        )}
        <li><strong>Track.</strong> Pick an activity tag such as Content Planning (written with the client tag, for example RAWCS Content Planning), describe the activity, click Start, then End. Click Add another session for the next task. Edit times in Sessions to add if you forgot to start or stop.</li>
        <li><strong>Save.</strong> Click Add hours to the file. Your work goes to the team log first, so it is safe even if the timesheet is busy. Save to team log only leaves the timesheet for a colleague to update.</li>
        <li><strong>Colleagues' work.</strong> Click Refresh (it flashes when there are updates). Tick entries and click Add selected to timesheet (placed in date order), Archive selected (hide, timesheet unchanged) or Delete selected (removes them, and their rows if they are in the timesheet).</li>
        <li><strong>Tools.</strong> Insert rows when Blank rows left is low. Save a copy for a backup. The bin icon deletes a row after you confirm, and its team log entry shows Removed from timesheet after Refresh. Undo restores the file as it was before your last change.</li>
      </ol>
      <p className="guide-tips">
        <button className="link dark" onClick={onManual}>Open the full user manual</button>{' '}
        Use Chrome or Edge. {m365 ? 'Close the timesheet in Excel before saving, because SharePoint holds files open for editing.' : 'Close the timesheet in Excel before saving.'} Dates are DD/MM/YYYY in Perth time.
      </p>
    </section>
  );
}

// Full user manual in a window over the app (same text as MANUAL.md).
export function Manual({ edition, onClose }) {
  const m = manual(edition);
  useEffect(() => {
    const onKey = (e) => e.key === 'Escape' && onClose();
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);
  return (
    <div className="modal" onClick={onClose}>
      <article className="card manual" role="dialog" aria-modal="true" aria-label="User manual" onClick={(e) => e.stopPropagation()}>
        <div className="manual-head">
          <h1>{m.title}</h1>
          <div className="row gap">
            <button className="ghost small" onClick={() => window.print()}>Print</button>
            <button className="small" onClick={onClose}>Close</button>
          </div>
        </div>
        <p className="manual-notice">{m.notice}</p>
        <nav className="manual-toc">{m.sections.map((s, i) => <a key={s.h} href={`#man-${i}`}>{s.h}</a>)}</nav>
        {m.sections.map((s, i) => (
          <section key={s.h} id={`man-${i}`}>
            <h2>{s.h}</h2>
            {s.blocks.map((b, j) => (
              <div key={j}>
                {b.p && <p>{b.p}</p>}
                {b.steps && <ol>{b.steps.map((t) => <li key={t}>{t}</li>)}</ol>}
                {b.list && <ul>{b.list.map((t) => <li key={t}>{t}</li>)}</ul>}
                {b.table && (
                  <table className="rows manual-table">
                    <thead><tr>{b.table[0].map((c) => <th key={c}>{c}</th>)}</tr></thead>
                    <tbody>{b.table.slice(1).map((r) => <tr key={r[0]}>{r.map((c, k) => <td key={k}>{k === 0 ? <strong>{c}</strong> : c}</td>)}</tr>)}</tbody>
                  </table>
                )}
              </div>
            ))}
          </section>
        ))}
      </article>
    </div>
  );
}
