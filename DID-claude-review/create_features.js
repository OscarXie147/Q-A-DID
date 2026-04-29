const XLSX = require('xlsx');

// 8个季度定义：按时间窗口严格划分
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

function parseDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    const d = new Date(String(val));
    return isNaN(d) ? null : d;
}

function charLength(str) {
    if (str == null) return 0;
    return String(str).length;
}

const wb = XLSX.readFile('./问答_合并商品信息.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v !== undefined ? String(c.v).trim() : 'col_' + col);
}

const timeRangeIdx = headers.indexOf('时间是否在范围');
const dateIdx = headers.indexOf('date');
const priceIdx = headers.indexOf('价格');
const questionIdx = headers.indexOf('question');
const answerIdx = headers.indexOf('answer');
const idIdx = 0;

console.log('timeRangeIdx:', timeRangeIdx, 'dateIdx:', dateIdx, 'priceIdx:', priceIdx, 'questionIdx:', questionIdx, 'answerIdx:', answerIdx);

// 收集每个商品ID的唯一价格（用于price_moderate）
const productPrices = {};
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const idCell = sheet[XLSX.utils.encode_cell({ r: row, c: idIdx })];
    const priceCell = sheet[XLSX.utils.encode_cell({ r: row, c: priceIdx })];
    const id = idCell && idCell.v !== undefined ? String(idCell.v) : null;
    const v = priceCell && priceCell.v !== undefined && priceCell.v !== null ? parseFloat(priceCell.v) : null;
    if (id && v !== null && !isNaN(v)) {
        productPrices[id] = v;
    }
}

const uniquePrices = [...new Set(Object.values(productPrices))].sort((a, b) => b - a);
const topThirdCount = Math.ceil(uniquePrices.length / 3);
const priceThreshold = uniquePrices[topThirdCount - 1];
console.log('\n=== price_moderate 计算 ===');
console.log('商品数:', Object.keys(productPrices).length, '去重价格数:', uniquePrices.length);
console.log('前1/3数量:', topThirdCount, '阈值:', priceThreshold);
console.log('阈值在去重数组中位置:', uniquePrices.indexOf(priceThreshold));
console.log('>=阈值的商品数:', Object.values(productPrices).filter(p => p >= priceThreshold).length);

// 构建新数据（去掉"时间是否在范围"列）
const filteredHeaders = headers.filter((_, i) => i !== timeRangeIdx);
const newCols = ['Treat', 'Quarter', 'Post', 'q_length', 'a_length', 'ln_q_length', 'ln_a_length', 'price_moderate'];
const allHeaders = [...filteredHeaders, ...newCols];

const newData = [allHeaders];
let inWindow = 0, outWindow = 0;
let q1=0,q2=0,q3=0,q4=0,q5=0,q6=0,q7=0,q8=0;
let pm1=0, pm0=0;

for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const rangeCell = sheet[XLSX.utils.encode_cell({ r: row, c: timeRangeIdx })];
    const dateCell = sheet[XLSX.utils.encode_cell({ r: row, c: dateIdx })];
    const priceCell = sheet[XLSX.utils.encode_cell({ r: row, c: priceIdx })];
    const idCell = sheet[XLSX.utils.encode_cell({ r: row, c: idIdx })];
    const qCell = sheet[XLSX.utils.encode_cell({ r: row, c: questionIdx })];
    const aCell = sheet[XLSX.utils.encode_cell({ r: row, c: answerIdx })];
    const id = idCell && idCell.v !== undefined ? String(idCell.v) : null;

    const inRange = rangeCell && rangeCell.v === true;
    if (!inRange) { outWindow++; continue; }
    inWindow++;

    // 原始行数据（跳过时间是否在范围列）
    const rowData = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
        if (col === timeRangeIdx) continue;
        const c = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
        rowData.push(c && c.v !== undefined ? c.v : null);
    }

    const dt = parseDate(dateCell && dateCell.v !== undefined ? dateCell.v : null);
    const qInfo = getQuarterInfo(dt);
    const price = priceCell && priceCell.v !== undefined && priceCell.v !== null ? parseFloat(priceCell.v) : null;
    const qLen = charLength(qCell && qCell.v !== undefined ? qCell.v : null);
    const aLen = charLength(aCell && aCell.v !== undefined ? aCell.v : null);

    const treat = 0;
    const quarter = qInfo ? qInfo.q : null;
    const post = qInfo ? qInfo.period : null;
    const priceModerate = price !== null && price >= priceThreshold ? 1 : 0;

    if (quarter === 1) q1++; else if (quarter === 2) q2++; else if (quarter === 3) q3++; else if (quarter === 4) q4++;
    else if (quarter === 5) q5++; else if (quarter === 6) q6++; else if (quarter === 7) q7++; else if (quarter === 8) q8++;
    if (priceModerate === 1) pm1++; else pm0++;

    newData.push([...rowData, treat, quarter, post, qLen, aLen, qLen > 0 ? Math.log(qLen) : null, aLen > 0 ? Math.log(aLen) : null, priceModerate]);
}

console.log('\n=== 过滤结果 ===');
console.log('窗口内:', inWindow, '窗口外:', outWindow);
console.log('季度分布: Q1=' + q1 + ' Q2=' + q2 + ' Q3=' + q3 + ' Q4=' + q4 + ' Q5=' + q5 + ' Q6=' + q6 + ' Q7=' + q7 + ' Q8=' + q8);
console.log('price_moderate: 1=' + pm1 + ' 0=' + pm0);

const newWb = XLSX.utils.book_new();
const newWs = XLSX.utils.aoa_to_sheet(newData);
XLSX.utils.book_append_sheet(newWb, newWs, '清洗后');
XLSX.writeFile(newWb, './问答_合并商品信息_清洗后.xlsx');
console.log('\n输出: 问答_合并商品信息_清洗后.xlsx');
console.log('总行数:', newData.length, '总列数:', allHeaders.length);
console.log('新表头:', allHeaders.join(' | '));
