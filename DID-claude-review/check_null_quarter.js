const XLSX = require('xlsx');

const wb = XLSX.readFile('./问答_合并商品信息_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v !== undefined ? String(c.v).trim() : 'col_' + col);
}
const quarterIdx = headers.indexOf('Quarter');
const postIdx = headers.indexOf('Post');
const dateIdx = headers.indexOf('date');
console.log('Quarter列:', quarterIdx, 'Post列:', postIdx, 'date列:', dateIdx);

const nullRows = [];
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const qAddr = XLSX.utils.encode_cell({ r: row, c: quarterIdx });
    const pAddr = XLSX.utils.encode_cell({ r: row, c: postIdx });
    const dAddr = XLSX.utils.encode_cell({ r: row, c: dateIdx });
    const qCell = sheet[qAddr];
    const pCell = sheet[pAddr];
    const dCell = sheet[dAddr];
    const qVal = qCell && qCell.v !== undefined ? qCell.v : null;
    const pVal = pCell && pCell.v !== undefined ? pCell.v : null;
    if (qVal === null || pVal === null) {
        nullRows.push({ row: row + 1, date: dCell && dCell.v !== undefined ? dCell.v : 'null', quarter: qVal, post: pVal });
    }
}

console.log('Quarter/Post为空的行数:', nullRows.length);
console.log('\n前20个为空的行:');
for (const r of nullRows.slice(0, 20)) {
    console.log('  Row' + r.row + ' date=' + r.date + ' quarter=' + r.quarter + ' post=' + r.post);
}

// 也统计有值的行
let hasQuarter = 0, hasPost = 0;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const qAddr = XLSX.utils.encode_cell({ r: row, c: quarterIdx });
    const pAddr = XLSX.utils.encode_cell({ r: row, c: postIdx });
    if (sheet[qAddr] && sheet[qAddr].v !== undefined && sheet[qAddr].v !== null) hasQuarter++;
    if (sheet[pAddr] && sheet[pAddr].v !== undefined && sheet[pAddr].v !== null) hasPost++;
}
console.log('\n有Quarter值的行:', hasQuarter, '有Post值的行:', hasPost);
