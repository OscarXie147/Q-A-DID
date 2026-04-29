const XLSX = require('xlsx');
const fs = require('fs');

const path = './商品详情_合并后.xlsx';
const wb = XLSX.readFile(path, { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : `col_${col}`);
}

let result = `总行数（含表头）: ${range.e.r - range.s.r + 1}\n`;
result += `列数: ${headers.length}\n\n`;
result += '表头: ' + headers.join(' | ') + '\n\n';
result += '前5行数据:\n';
for (let row = range.s.r + 1; row <= Math.min(range.e.r, range.s.r + 5); row++) {
    const vals = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
        const c = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
        vals.push(c && c.v != null ? String(c.v) : '(空)');
    }
    result += `  Row${row}: ${vals.join(' | ')}\n`;
}

const idSet = new Set();
let duplicate = false;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const c = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
    const id = c && c.v != null ? String(c.v) : 'null';
    if (idSet.has(id)) { duplicate = true; break; }
    idSet.add(id);
}
result += `\nID唯一性检查: ${duplicate ? '有重复!' : '全部唯一，OK'}\n`;
result += `总行数（不含表头）: ${idSet.size}\n`;

fs.writeFileSync('./merge_verify.txt', result, 'utf-8');
console.log(result);
