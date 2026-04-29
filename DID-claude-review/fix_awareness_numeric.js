const XLSX = require('xlsx');

const wb = XLSX.readFile('./商品详情_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];

const headers = [];
const range = XLSX.utils.decode_range(sheet['!ref']);
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : 'col_' + col);
}
const idx = headers.indexOf('brand_awareness');

let changed = 0;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const addr = XLSX.utils.encode_cell({ r: row, c: idx });
    const cell = sheet[addr];
    if (cell && typeof cell.v === 'string') {
        const num = parseInt(cell.v, 10);
        cell.v = num;
        cell.t = 'n';  // set type to number
        delete cell.w;  // remove formatted string
        delete cell.h;  // remove html representation
        changed++;
    }
}

console.log('转换了', changed, '个单元格为数值格式');

// 验证
const sampleAddr = XLSX.utils.encode_cell({ r: 1, c: idx });
const sample = sheet[sampleAddr];
console.log('验证Row2: t=' + sample.t + ' v=' + sample.v + ' typeof=' + typeof sample.v);

XLSX.writeFile(wb, './商品详情_清洗后.xlsx');
console.log('完成，已保存！');
