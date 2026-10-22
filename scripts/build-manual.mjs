// Writes MANUAL.md from src/manual.js. Usage: node scripts/build-manual.mjs local|m365 [out]
import fs from 'node:fs';
import { manualMarkdown } from '../src/manual.js';
const [edition = 'local', out = 'MANUAL.md'] = process.argv.slice(2);
fs.writeFileSync(out, manualMarkdown(edition));
console.log(`Wrote ${out} (${edition} edition)`);
