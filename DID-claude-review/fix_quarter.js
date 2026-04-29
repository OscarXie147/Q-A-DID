const XLSX = require('xlsx');

const QUARTERS = [
    { q: 1, start: [2022,10,1], end: [2022,12,31], period: 0 },
    { q: 2, start: [2023, 1, 1], end: [2023, 3,31], period: 0 },
    { q: 3, start: [2023, 4, 1], end: [2023, 6,30], period: 0 },
    { q: 4, start: [2023, 7, 1], end: [2023, 9,30], period: 0 },
    { q: 5, start: [2023,10, 1], end: [2023,12,31], period: 1 },
    { q: 6, start: [2024, 1, 1], end: [2024, 3,31], period: 1 },
    { q: 7, start: [2024, 4, 1], end: [2024, 6,30], period: 1 },
    { q: 8, start: [2024, 7, 1], end: [2024, 9,30], period: 1 },
];

function parseDateParts(val) {
    if (!val) return null;
    if (val instanceof Date) return [val.getFullYear(), val.getMonth()+1, val.getDate()];
    const s = String(val).trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return [parseInt(m[1]), parseInt(m[2]), parseInt(m[3])];
    const d = new Date(s);
    if (isNaN(d)) return null;
    return [d.getFullYear(), d.getMonth()+1, d.getDate()];
}

function compareDateParts(a, b) {
    if (a[0] !== b[0]) return a[0] - b[0];
    if (a[1] !== b[1]) return a[1] - b[1];
    return a[2] - b[2];
}

function getQuarterInfo(val) {
    const parts = parseDateParts(val);
    if (!parts) return null;
    for (const q of QUARTERS) {
        if (compareDateParts(parts, q.start) >= 0 && compareDateParts(parts, q.end) <= 0) return q;
    }
    return null;
}

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
const postIdx = headers.indexOf('Post');
console.log('dateIdx:', dateIdx, 'quarterIdx:', quarterIdx, 'postIdx:', postIdx);

let fixed = 0, alreadySet = 0, stillNull = 0;

for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const dAddr = XLSX.utils.encode_cell({ r: row, c: dateIdx });
    const qAddr = XLSX.utils.encode_cell({ r: row, c: quarterIdx });
    const pAddr = XLSX.utils.encode_cell({ r: row, c: postIdx });
    const dCell = sheet[dAddr];
    const qCell = sheet[qAddr];

    const qVal = qCell && qCell.v !== undefined ? qCell.v : null;
    if (qVal === null || qVal === undefined) {
        const dateVal = dCell && dCell.v !== undefined ? dCell.v : null;
        const info = getQuarterInfo(dateVal);
        if (info) {
            if (!sheet[qAddr]) sheet[qAddr] = { t: 'n' };
            if (!sheet[pAddr]) sheet[pAddr] = { t: 'n' };
            sheet[qAddr].v = info.q;
            sheet[pAddr].v = info.period;
            fixed++;
        } else {
            stillNull++;
        }
    } else {
        alreadySet++;
    }
}

console.log('已设置过:', alreadySet, '本次修复:', fixed, '仍为null:', stillNull);

// 验证
let nullCount = 0;
const qStats = {};
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const qAddr = XLSX.utils.encode_cell({ r: row, c: quarterIdx });
    const qCell = sheet[qAddr];
    if (!qCell || qCell.v === null || qCell.v === undefined) nullCount++;
    else qStats[qCell.v] = (qStats[qCell.v] || 0) + 1;
}
console.log('剩余null行:', nullCount);
console.log('季度分布:', qStats);

XLSX.writeFile(wb, './问答_合并商品信息_清洗后.xlsx');
console.log('已保存！');
