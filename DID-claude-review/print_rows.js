const XLSX = require('xlsx');

const path = './商品详情_合并后.xlsx';
const wb = XLSX.readFile(path, { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : `col_${col}`);
}

console.log('表头:', headers.join(' | '));
console.log('\n数据行（前10行）:');
for (let row = range.s.r + 1; row <= Math.min(range.e.r, range.s.r + 10); row++) {
    const vals = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
        const c = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
        const v = c && c.v != null ? String(c.v) : '(空)';
        vals.push(v.length > 60 ? v.slice(0, 60) + '...' : v);
    }
    console.log(`\n--- Row ${row} ---`);
    for (let i = 0; i < headers.length; i++) {
        console.log(`  [${headers[i]}]: ${vals[i]}`);
    }
}
