// One-off: correct swapped or text dates in a timesheet's Date Log column, editing only those cells.
// Usage: node scripts/fix-dates.mjs in.xlsx out.xlsx "27=2026-08-12" "88=2026-10-01" ...
import fs from 'node:fs';
import { readWorkbook, analyse } from '../src/core/timesheet.js';
import { internal } from '../src/core/xlsxPatch.js';

const [src, out, ...pairs] = process.argv.slice(2);
const bytes = fs.readFileSync(src);
const info = analyse(await readWorkbook(bytes));
const { openPackage, savePackage, Sheet, Styles, excelSerial } = internal;
const pkg = await openPackage(bytes, info.sheetName);
const sheet = new Sheet(pkg.sheetXml);
const styles = new Styles(pkg.stylesXml);
const c = info.cols.date;
for (const p of pairs) {
  const [r, d] = p.split('=');
  sheet.put(Number(r), c, { kind: 'number', value: excelSerial(d, pkg.date1904), style: sheet.styleOf(Number(r), c) });
}
// Every date cell in the data area shows DD/MM/YYYY.
for (let r = info.firstDataRow; r < (info.totalRow || info.lastUsedRow + 1); r++) {
  const s = sheet.styleOf(r, c);
  if (s !== undefined) { const cell = sheet.cell(r, c); const t = cell.attrs.find(([k]) => k === 't')?.[1]; sheet.put(r, c, { kind: 'raw', rawType: t, rawInner: cell.inner, style: styles.auDateStyle(s) }); }
}
fs.writeFileSync(out, Buffer.from(await savePackage(pkg, sheet, styles)));
