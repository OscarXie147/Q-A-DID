const XLSX = require('xlsx');

const wb = XLSX.readFile('./商品详情_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : 'col_' + col);
}
const brandIdx = headers.indexOf('品牌');
const baiduIdx = headers.indexOf('百度指数');
console.log('品牌列:', brandIdx, '百度指数列:', baiduIdx);
console.log('表头:', headers);

const emptyRows = [];
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const brandCell = sheet[XLSX.utils.encode_cell({ r: row, c: brandIdx })];
    const indexCell = sheet[XLSX.utils.encode_cell({ r: row, c: baiduIdx })];
    const brand = brandCell && brandCell.v != null ? String(brandCell.v) : '';
    const indexVal = indexCell && indexCell.v != null ? indexCell.v : null;
    if (indexVal === null || indexVal === '') {
        emptyRows.push({ row: row + 1, brand });
    }
}

console.log('\n百度指数仍为空的行数:', emptyRows.length);
for (const r of emptyRows) {
    console.log('  Row' + r.row + ': ' + r.brand);
}
