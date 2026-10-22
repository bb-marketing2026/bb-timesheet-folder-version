// Run: npm test -- "path/to/timesheet.xlsx"
// Appends sample rows, deletes one, and checks that nothing else in the file changed.
import fs from 'node:fs';
import JSZip from 'jszip';
import { readWorkbook, analyse } from '../src/core/timesheet.js';
import { appendSessionsSafe, deleteEntrySafe, insertRowsSafe } from '../src/core/xlsxPatch.js';

const file = process.argv[2];
const out = process.argv[3] || 'test-output.xlsx';
const original = fs.readFileSync(file);
const show = async (bytes, label) => {
  const info = analyse(await readWorkbook(bytes));
  console.log(`${label}: last row ${info.lastUsedRow}, total ${info.total}, rows left ${info.rowsLeft}`);
  info.entries.slice(-5).forEach((e) => console.log(`   ${e.row} ${e.date}${e.hasDate ? '*' : ' '} ${e.tag} | ${e.activity} | ${e.hours}`));
  return info;
};
const before = await show(original, 'Before');

const s = (date, tag, activity, from, to) => ({ date, tag, activity, start: `${date}T${from}:00+08:00`, end: `${date}T${to}:00+08:00` });
const added = await appendSessionsSafe(original, before.sheetName, [
  s('2026-10-22', 'RAWCS Content Hub', 'Updated hub images', '09:00', '09:45'),
  s('2026-10-22', 'RAWCS Content Hub', 'Fixed captions & <tags>', '10:00', '10:20'),
  s('2026-10-22', 'RAWCS Paid Advertisement', 'Checked Google Ads spend', '11:00', '11:10'),
  s('2026-10-23', 'RAWCS Social Media Management', 'Queued posts', '08:00', '08:30'),
], { combine: true });
console.log('Wrote rows', added.rows);
const a = await show(added.bytes, 'After add');

// Delete the first row of 22/10 (it carries the date), the next 22/10 row should inherit it.
const target = a.entries.find((e) => e.row === added.rows[0]);
const del = await deleteEntrySafe(added.bytes, a.sheetName, target);
const d = await show(del.bytes, `After deleting row ${target.row}`);

// Stale delete must be refused.
try { await deleteEntrySafe(del.bytes, a.sheetName, target); console.log('ERROR: stale delete allowed'); }
catch (e) { console.log('Stale delete refused:', e.message.slice(0, 60)); }

// Insert 10 rows above Total Hours.
const ins = await insertRowsSafe(del.bytes, a.sheetName, 10);
console.log('Inserted 10 rows: Total Hours now row', ins.totalRow, '| rows left', ins.rowsLeft);

// Every part other than the sheet, styles, workbook calc flag and app.xml must be byte-identical.
const z0 = await JSZip.loadAsync(original);
const z1 = await JSZip.loadAsync(ins.bytes);
const n0 = Object.keys(z0.files).filter((n) => !z0.files[n].dir);
const n1 = Object.keys(z1.files).filter((n) => !z1.files[n].dir);
console.log('Parts before/after:', n0.length, n1.length, n0.join() === n1.join() ? '(same order)' : '(DIFFERENT)');
const changed = [];
for (const n of n0) {
  const [x, y] = await Promise.all([z0.file(n).async('uint8array'), z1.file(n)?.async('uint8array')]);
  if (!y || Buffer.compare(Buffer.from(x), Buffer.from(y)) !== 0) changed.push(n);
}
console.log('Changed parts:', changed);
fs.writeFileSync(out, Buffer.from(ins.bytes));
fs.writeFileSync(out.replace('.xlsx', '-added.xlsx'), Buffer.from(added.bytes));
