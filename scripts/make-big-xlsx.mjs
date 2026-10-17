#!/usr/bin/env node
// Writes a large .xlsx for the manual "100 MB file" check, streaming so memory stays flat.
//   npm run make-big-xlsx -- [sizeMB=100] [out=big.xlsx]
// Upload the result from the demo site, then open it with Edit in Excel.
import { createRequire } from 'node:module';
import { randomBytes } from 'node:crypto';
import { statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(join(root, 'apps/demo-site/package.json'));
const ExcelJS = require('exceljs');

const targetMb = Number(process.argv[2] ?? 100);
const out = resolve(process.argv[3] ?? `big-${targetMb}mb.xlsx`);

// Random base64 compresses to about 3/4, so the zipped size tracks the rows
// written. ROWS_PER_MB was measured on this generator.
const ROWS_PER_MB = 1290;
const rows = Math.ceil(targetMb * ROWS_PER_MB);
const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: out, useSharedStrings: false, useStyles: false });
const sheet = workbook.addWorksheet('Data');
sheet.addRow(['#', 'A', 'B', 'C', 'D']).commit();
for (let i = 1; i <= rows; i++) {
  sheet.addRow([i, ...Array.from({ length: 4 }, () => randomBytes(192).toString('base64'))]).commit();
  if (i % 20000 === 0) process.stdout.write(`\r${Math.round((i / rows) * 100)}%`);
}
await workbook.commit();
console.log(`\rWrote ${out} (${(statSync(out).size / 1024 / 1024).toFixed(1)} MB, ${rows} rows)`);
