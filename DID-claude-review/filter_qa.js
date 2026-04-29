const XLSX = require('xlsx');

const qaPath = './问答.xlsx';
const outPath = './问答_符合条件.xlsx';

const PRE_QUARTERS = [
    ['pre_Q1', new Date('2022-10-01'), new Date('2022-12-31 23:59:59')],
    ['pre_Q2', new Date('2023-01-01'), new Date('2023-03-31 23:59:59')],
    ['pre_Q3', new Date('2023-04-01'), new Date('2023-06-30 23:59:59')],
    ['pre_Q4', new Date('2023-07-01'), new Date('2023-09-30 23:59:59')],
];
const POST_QUARTERS = [
    ['post_Q1', new Date('2023-10-01'), new Date('2023-12-31 23:59:59')],
    ['post_Q2', new Date('2024-01-01'), new Date('2024-03-31 23:59:59')],
    ['post_Q3', new Date('2024-04-01'), new Date('2024-06-30 23:59:59')],
    ['post_Q4', new Date('2024-07-01'), new Date('2024-09-30 23:59:59')],
];
const ALL_QUARTERS = [...PRE_QUARTERS, ...POST_QUARTERS];

function parseDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    const s = String(val).trim();
    const d = new Date(s);
    return isNaN(d) ? null : d;
}

function getQuarter(dt) {
    if (!dt) return null;
    for (const q of ALL_QUARTERS) {
        if (dt >= q[1] && dt <= q[2]) return q[0];
    }
    return null;
}

const wb = XLSX.readFile(qaPath, { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v).trim() : 'col_' + col);
}
const idColIdx = headers.indexOf('auctionNumId');
const dateColIdx = headers.indexOf('date');
console.log('ID列:', idColIdx, '日期列:', dateColIdx);

// 统计每个商品各季度记录数
const productQuarters = {};
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const idCellAddr = XLSX.utils.encode_cell({ r: row, c: idColIdx });
    const dateCellAddr = XLSX.utils.encode_cell({ r: row, c: dateColIdx });
    const idCell = sheet[idCellAddr];
    const dateCell = sheet[dateCellAddr];
    const id = idCell && idCell.v != null ? String(idCell.v) : null;
    const dt = parseDate(dateCell && dateCell.v != null ? dateCell.v : null);
    if (!id || !dt) continue;

    const q = getQuarter(dt);
    if (!q) continue;

    if (!productQuarters[id]) productQuarters[id] = {};
    productQuarters[id][q] = (productQuarters[id][q] || 0) + 1;
}

// 筛选8个季度每季>=1
const qualifiedIds = new Set();
for (const [id, quarters] of Object.entries(productQuarters)) {
    let ok = true;
    for (const q of ALL_QUARTERS) {
        if ((quarters[q[0]] || 0) < 1) { ok = false; break; }
    }
    if (ok) qualifiedIds.add(id);
}

console.log('总商品数:', Object.keys(productQuarters).length);
console.log('符合条件商品数:', qualifiedIds.size);

// 提取数据行
const filteredData = [headers];
let filteredCount = 0;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const idCellAddr = XLSX.utils.encode_cell({ r: row, c: idColIdx });
    const idCell = sheet[idCellAddr];
    const id = idCell && idCell.v != null ? String(idCell.v) : null;
    if (id && qualifiedIds.has(id)) {
        const rowData = [];
        for (let col = range.s.c; col <= range.e.c; col++) {
            const c = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
            rowData.push(c && c.v != null ? c.v : null);
        }
        filteredData.push(rowData);
        filteredCount++;
    }
}

const newWb = XLSX.utils.book_new();
const newWs = XLSX.utils.aoa_to_sheet(filteredData);
XLSX.utils.book_append_sheet(newWb, newWs, '符合条件问答');
XLSX.writeFile(newWb, outPath);

console.log('\n筛选出', filteredCount, '条问答记录');
console.log('唯一商品数:', qualifiedIds.size);
console.log('输出文件:', outPath);

const quarterTotals = {};
for (const q of ALL_QUARTERS) quarterTotals[q[0]] = 0;
for (const [id, quarters] of Object.entries(productQuarters)) {
    if (qualifiedIds.has(id)) {
        for (const [label, cnt] of Object.entries(quarters)) {
            quarterTotals[label] = (quarterTotals[label] || 0) + cnt;
        }
    }
}
console.log('\n各季度记录数:');
for (const q of ALL_QUARTERS) console.log(' ', q[0], ':', quarterTotals[q[0]]);
