const XLSX = require('xlsx');

const QUARTERS = [
    { q: 1, start: new Date('2022-10-01'), end: new Date('2022-12-31'), period: 0 },
    { q: 2, start: new Date('2023-01-01'), end: new Date('2023-03-31'), period: 0 },
    { q: 3, start: new Date('2023-04-01'), end: new Date('2023-06-30'), period: 0 },
    { q: 4, start: new Date('2023-07-01'), end: new Date('2023-09-30'), period: 0 },
    { q: 5, start: new Date('2023-10-01'), end: new Date('2023-12-31'), period: 1 },
    { q: 6, start: new Date('2024-01-01'), end: new Date('2024-03-31'), period: 1 },
    { q: 7, start: new Date('2024-04-01'), end: new Date('2024-06-30'), period: 1 },
    { q: 8, start: new Date('2024-07-01'), end: new Date('2024-09-30'), period: 1 },
];

function getQuarterInfo(dt) {
    if (!dt) return null;
    const d = dt instanceof Date ? dt : new Date(dt);
    if (isNaN(d)) return null;
    for (const q of QUARTERS) {
        if (d >= q.start && d <= q.end) return q;
    }
    return null;
}

// 检查几个边界日期
const testDates = ['2024-09-29', '2024-09-30', '2023-03-31', '2023-06-30', '2023-09-30', '2023-03-31 11:49:18', '2022-12-31'];
for (const s of testDates) {
    const d = new Date(s);
    const info = getQuarterInfo(d);
    console.log(s + ' -> ' + d.toISOString() + ' quarter=' + (info ? info.q : 'null'));
}

// 检查xlsx文件中的实际date值类型
const wb = XLSX.readFile('./问答_合并商品信息_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v !== undefined ? String(c.v).trim() : 'col_' + col);
}
const dateIdx = headers.indexOf('date');
const quarterIdx = headers.indexOf('Quarter');

console.log('\n检查前10个Quarter为空的行:');
let count = 0;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const qAddr = XLSX.utils.encode_cell({ r: row, c: quarterIdx });
    const dAddr = XLSX.utils.encode_cell({ r: row, c: dateIdx });
    const qCell = sheet[qAddr];
    const dCell = sheet[dAddr];
    if ((!qCell || qCell.v === null || qCell.v === undefined) && count < 10) {
        const rawDate = dCell && dCell.v !== undefined ? dCell.v : 'null';
        const dateObj = dCell && dCell.v !== undefined ? new Date(dCell.v) : null;
        const parsed = getQuarterInfo(dateObj);
        console.log('Row' + (row+1) + ': date t=' + (dCell ? dCell.t : 'null') + ' v=' + rawDate + ' parsed=' + (parsed ? parsed.q : 'null'));
        count++;
    }
}
