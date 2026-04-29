const XLSX = require('xlsx');

const wb = XLSX.readFile('./商品详情_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : 'col_' + col);
}
const brandIdx = headers.indexOf('品牌');
const baiduIdx = headers.indexOf('百度指数');
const awarenessIdx = headers.indexOf('brand_awareness');

let fixed = 0;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const biAddr = XLSX.utils.encode_cell({ r: row, c: baiduIdx });
    const aiAddr = XLSX.utils.encode_cell({ r: row, c: awarenessIdx });
    const brAddr = XLSX.utils.encode_cell({ r: row, c: brandIdx });
    const biVal = sheet[biAddr].v;
    const aiVal = sheet[aiAddr].v;

    if (biVal === null && aiVal === null) {
        sheet[aiAddr].v = 0;
        const brand = sheet[brAddr].v;
        console.log('修复 Row' + (row+1) + ' (' + brand + '): awareness null -> 0');
        fixed++;
    }
}

const stats = { '1': 0, '0': 0, 'null': 0 };
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const aiAddr = XLSX.utils.encode_cell({ r: row, c: awarenessIdx });
    const v = sheet[aiAddr].v;
    if (v === 1) stats['1']++;
    else if (v === 0) stats['0']++;
    else stats['null']++;
}
console.log('\n修复后统计: 1=' + stats['1'] + ' 0=' + stats['0'] + ' null=' + stats['null']);

XLSX.writeFile(wb, './商品详情_清洗后.xlsx');
console.log('完成，已保存！修复', fixed, '行');
