const XLSX = require('xlsx');

const wb = XLSX.readFile('./问答_合并商品信息_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v !== undefined ? String(c.v).trim() : 'col_' + col);
}
const qLenIdx = headers.indexOf('q_length');
const aLenIdx = headers.indexOf('a_length');
const lnQIdx = headers.indexOf('ln_q_length');
const lnAIdx = headers.indexOf('ln_a_length');
console.log('q_length:', qLenIdx, 'a_length:', aLenIdx, 'ln_q_length:', lnQIdx, 'ln_a_length:', lnAIdx);

// 检查ln_a_length = 0 的情况（对应 a_length = 1）
let aLen1 = 0, lnA0 = 0;
let qLen0 = 0, lnQ0 = 0;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const aAddr = XLSX.utils.encode_cell({ r: row, c: aLenIdx });
    const qAddr = XLSX.utils.encode_cell({ r: row, c: qLenIdx });
    const lnAAddr = XLSX.utils.encode_cell({ r: row, c: lnAIdx });
    const lnQAddr = XLSX.utils.encode_cell({ r: row, c: lnQIdx });
    const aCell = sheet[aAddr];
    const qCell = sheet[qAddr];
    const lnACell = sheet[lnAAddr];
    const lnQCell = sheet[lnQAddr];
    const aVal = aCell && aCell.v !== undefined ? aCell.v : null;
    const qVal = qCell && qCell.v !== undefined ? qCell.v : null;
    const lnAVal = lnACell && lnACell.v !== undefined ? lnACell.v : null;
    const lnQVal = lnQCell && lnQCell.v !== undefined ? lnQCell.v : null;
    if (aVal === 1) aLen1++;
    if (lnAVal === 0) lnA0++;
    if (qVal === 1) qLen0++;
    if (lnQVal === 0) lnQ0++;
}
console.log('\na_length=1的行数:', aLen1);
console.log('ln_a_length=0的行数（因ln(1)=0）:', lnA0);
console.log('q_length=1的行数:', qLen0);
console.log('ln_q_length=0的行数（因ln(1)=0）:', lnQ0);

if (aLen1 > 0) console.log('\na_length存在为1的情况，需要修复ln_a_length');
if (qLen0 > 0) console.log('q_length存在为1的情况，需要修复ln_q_length');
