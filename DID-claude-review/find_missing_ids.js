const XLSX = require('xlsx');

const path = './商品详情_清洗后.xlsx';
const wb = XLSX.readFile(path, { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : `col_${col}`);
}

console.log('缺失商品的完整信息:\n');
for (const row of [23, 72]) {
    console.log(`--- 第${row}行（Excel行号${row+1}）---`);
    for (let col = range.s.c; col <= range.e.c; col++) {
        const c = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
        const v = c && c.v != null ? String(c.v) : '(空)';
        console.log(`  [${headers[col]}]: ${v}`);
    }
    console.log();
}
