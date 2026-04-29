const XLSX = require('xlsx');

const wb = XLSX.readFile('./问答_合并商品信息_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v !== undefined ? String(c.v).trim() : 'col_' + col);
}

const idCols = ['auctionNumId', 'questionId'];
for (const colName of idCols) {
    const colIdx = headers.indexOf(colName);
    if (colIdx === -1) { console.log('未找到列:', colName); continue; }
    let changed = 0;
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
        const addr = XLSX.utils.encode_cell({ r: row, c: colIdx });
        const cell = sheet[addr];
        if (cell && cell.v !== undefined && cell.v !== null) {
            cell.v = String(cell.v);
            cell.t = 's';
            changed++;
        }
    }
    console.log(colName + ' 已转为文本格式，修改了 ' + changed + ' 个单元格');
    const sample = sheet[XLSX.utils.encode_cell({ r: range.s.r + 1, c: colIdx })];
    console.log('示例: t=' + sample.t + ' v=' + sample.v + ' typeof=' + typeof sample.v);
}

XLSX.writeFile(wb, './问答_合并商品信息_清洗后.xlsx');
console.log('\n已保存！');
