const XLSX = require('xlsx');

const wb = XLSX.readFile('./问答_符合条件.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];

const headers = [];
const range = XLSX.utils.decode_range(sheet['!ref']);
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v).trim() : 'col_' + col);
}
const colIdx = headers.indexOf('时间是否在范围');
const dateColIdx = headers.indexOf('date');
console.log('列索引: colIdx=' + colIdx + ' dateColIdx=' + dateColIdx);

const values = {};
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const c = sheet[XLSX.utils.encode_cell({ r: row, c: colIdx })];
    const v = c && c.v !== undefined && c.v !== null ? String(c.v).trim() : 'null';
    values[v] = (values[v] || 0) + 1;
}
console.log('值分布:', values);

console.log('\n前10行:');
for (let row = range.s.r + 1; row <= range.s.r + 10; row++) {
    const dateAddr = XLSX.utils.encode_cell({ r: row, c: dateColIdx });
    const rangeAddr = XLSX.utils.encode_cell({ r: row, c: colIdx });
    const dateCell = sheet[dateAddr];
    const rangeCell = sheet[rangeAddr];
    console.log('  date=' + (dateCell && dateCell.v ? dateCell.v : '') + ' 时间是否在范围=' + (rangeCell && rangeCell.v !== undefined ? rangeCell.v : ''));
}
