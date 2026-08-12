// debug-workbook-v3.js
// Run with: node debug-workbook-v3.js "path\to\your\file.xlsx"
//
// Re-parses with WTF:true, which makes SheetJS throw on internal parse
// issues instead of silently skipping the sheet. This should reveal exactly
// why "data all" isn't making it into wb.Sheets.

const XLSX = require('xlsx');
const fs = require('fs');

const filePath = process.argv[2];
if (!filePath) {
  console.error('Usage: node debug-workbook-v3.js "path\\to\\file.xlsx"');
  process.exit(1);
}

const buf = fs.readFileSync(filePath);

console.log('--- Attempt 1: normal read (baseline) ---');
try {
  const wb1 = XLSX.read(buf, { type: 'buffer', cellDates: true });
  console.log('OK. "data all" present in Sheets:', wb1.Sheets['data all'] !== undefined);
} catch (e) {
  console.log('Threw:', e && e.message);
}

console.log('\n--- Attempt 2: WTF:true (surface internal parse errors) ---');
try {
  const wb2 = XLSX.read(buf, { type: 'buffer', cellDates: true, WTF: true });
  console.log('OK (no error thrown). "data all" present in Sheets:', wb2.Sheets['data all'] !== undefined);
} catch (e) {
  console.log('THREW AN ERROR (this is what we want to see):');
  console.log(e && e.stack ? e.stack : e);
}

console.log('\n--- Attempt 3: dense mode ---');
try {
  const wb3 = XLSX.read(buf, { type: 'buffer', cellDates: true, dense: true });
  console.log('OK. "data all" present in Sheets:', wb3.Sheets['data all'] !== undefined);
} catch (e) {
  console.log('Threw:', e && e.message);
}

console.log('\n--- Attempt 4: sheets option limited to just "data all" by index ---');
try {
  // First find the index of "data all" from a normal read's SheetNames
  const wbNames = XLSX.read(buf, { type: 'buffer', bookSheets: true });
  const targetIdx = wbNames.SheetNames.findIndex((n) => n.trim().toLowerCase() === 'data all');
  console.log('Index of "data all":', targetIdx);
  const wb4 = XLSX.read(buf, { type: 'buffer', cellDates: true, sheets: [targetIdx] });
  console.log('SheetNames returned:', wb4.SheetNames);
  console.log('Sheets keys returned:', Object.keys(wb4.Sheets));
} catch (e) {
  console.log('THREW:', e && e.stack ? e.stack : e);
}

console.log('\n--- Node / SheetJS versions ---');
console.log('Node:', process.version);
try {
  console.log('xlsx (SheetJS) version:', require('xlsx/package.json').version);
} catch (e) {
  console.log('(could not read xlsx package.json)');
}