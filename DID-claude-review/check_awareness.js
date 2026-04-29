const XLSX = require('xlsx');

const wb = XLSX.readFile('./商品详情_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : 'col_' + col);
}
const awarenessIdx = headers.indexOf('brand_awareness');

console.log('brand_awareness列所有值:');
const stats = {};
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const ai = sheet[XLSX.utils.encode_cell({ r: row, c: awarenessIdx })];
    const v = ai ? ai.v : 'undefined';
    const k = String(v);
    stats[k] = (stats[k] || 0) + 1;
    if (row <= range.s.r + 5) {
        console.log('  Row' + (row+1) + ': t=' + typeof v + ' v=' + JSON.stringify(v));
    }
}
console.log('\n值分布:', stats);
