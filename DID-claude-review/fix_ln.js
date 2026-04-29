const XLSX = require('xlsx');

const wb = XLSX.readFile('./问答_合并商品信息_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v !== undefined ? String(c.v).trim() : 'col_' + col);
}
const aLenIdx = headers.indexOf('a_length');
const lnAIdx = headers.indexOf('ln_a_length');
console.log('a_length:', aLenIdx, 'ln_a_length:', lnAIdx);

let fixed = 0;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const aAddr = XLSX.utils.encode_cell({ r: row, c: aLenIdx });
    const lnAAddr = XLSX.utils.encode_cell({ r: row, c: lnAIdx });
    const aCell = sheet[aAddr];
    const lnACell = sheet[lnAAddr];
    const aVal = aCell && aCell.v !== undefined && aCell.v !== null ? parseFloat(aCell.v) : null;
    if (aVal !== null && !isNaN(aVal)) {
        const newLnA = Math.log(aVal + 1);
        if (lnACell && lnACell.v !== undefined) {
            sheet[lnAAddr].v = newLnA;
        }
        fixed++;
    }
}

console.log('修复了', fixed, '行的ln_a_length');

// 验证：原来ln_a=0的行现在是ln(1+1)=ln(2)
let stillZero = 0, checkCount = 0;
for (let row = range.s.r + 1; row <= range.e.r && checkCount < 5; row++) {
    const aAddr = XLSX.utils.encode_cell({ r: row, c: aLenIdx });
    const lnAAddr = XLSX.utils.encode_cell({ r: row, c: lnAIdx });
    const aCell = sheet[aAddr];
    const lnACell = sheet[lnAAddr];
    const aVal = aCell && aCell.v !== undefined ? parseFloat(aCell.v) : null;
    const lnAVal = lnACell && lnACell.v !== undefined ? lnACell.v : null;
    if (aVal === 1) {
        console.log('a_length=1 -> ln_a_length=' + lnAVal + ' (应为ln(2)=' + Math.log(2) + ')');
        checkCount++;
    }
    if (lnAVal === 0) stillZero++;
}
console.log('ln_a_length仍为0的行数:', stillZero);

XLSX.writeFile(wb, './问答_合并商品信息_清洗后.xlsx');
console.log('已保存！');
