const XLSX = require('xlsx');

const path = './商品详情_清洗后.xlsx';
const wb = XLSX.readFile(path, { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : `col_${col}`);
}

const colMissing = {};
for (const h of headers) colMissing[h] = [];

for (let row = range.s.r + 1; row <= range.e.r; row++) {
    for (let col = range.s.c; col <= range.e.c; col++) {
        const c = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
        const v = c && c.v != null ? c.v : null;
        if (v === null || v === '' || v === undefined) {
            colMissing[headers[col]].push(row);
        }
    }
}

let result = '';
result += `总行数（不含表头）: ${range.e.r - range.s.r}\n\n`;
result += '各列缺失值统计:\n';
result += '================================================\n';

let totalMissing = 0;
for (const h of headers) {
    const missing = colMissing[h];
    if (missing.length > 0) {
        result += `\n【${h}】缺失 ${missing.length} 行\n`;
        result += `  行号: ${missing.join(', ')}\n`;
        totalMissing += missing.length;
    }
}

if (totalMissing === 0) {
    result += '\n无缺失值！\n';
} else {
    result += `\n\n总计缺失值: ${totalMissing} 个\n`;
}

console.log(result);
