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
const lnQIdx = headers.indexOf('ln_q_length');
console.log('q_length:', qLenIdx, 'ln_q_length:', lnQIdx);

let fixed = 0;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const qAddr = XLSX.utils.encode_cell({ r: row, c: qLenIdx });
    const lnQAddr = XLSX.utils.encode_cell({ r: row, c: lnQIdx });
    const qCell = sheet[qAddr];
    const lnQCell = sheet[lnQAddr];
    const qVal = qCell && qCell.v !== undefined && qCell.v !== null ? parseFloat(qCell.v) : null;
    if (qVal !== null && !isNaN(qVal)) {
        const newLnQ = Math.log(qVal + 1);
        if (lnQCell && lnQCell.v !== undefined) {
            sheet[lnQAddr].v = newLnQ;
        }
        fixed++;
    }
}

console.log('修复了', fixed, '行的ln_q_length');

XLSX.writeFile(wb, './问答_合并商品信息_清洗后.xlsx');
console.log('已保存！');
