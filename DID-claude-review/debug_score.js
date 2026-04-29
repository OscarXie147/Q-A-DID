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

const scoreIdx = headers.indexOf('店铺评分');
console.log('店铺评分列索引:', scoreIdx);

// 打印前3行的原始值类型和内容
for (let row = range.s.r + 1; row <= range.s.r + 4; row++) {
    const c = sheet[XLSX.utils.encode_cell({ r: row, c: scoreIdx })];
    console.log(`\nRow${row} 店铺评分:`);
    console.log('  typeof:', typeof c.v);
    console.log('  isArray:', Array.isArray(c.v));
    console.log('  value:', JSON.stringify(c.v));
}
