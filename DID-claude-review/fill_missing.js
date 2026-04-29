const XLSX = require('xlsx');

const path = './商品详情_清洗后.xlsx';
const wb = XLSX.readFile(path, { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : 'col_' + col);
}

const colIdx = {};
for (let i = 0; i < headers.length; i++) colIdx[headers[i]] = i;

// 计算各列平均值
const sums = { 宝贝质量: 0, 物流速度: 0, 服务保障: 0 };
const counts = { 宝贝质量: 0, 物流速度: 0, 服务保障: 0 };

for (let row = range.s.r + 1; row <= range.e.r; row++) {
    for (const col of ['宝贝质量', '物流速度', '服务保障']) {
        const cellAddr = XLSX.utils.encode_cell({ r: row, c: colIdx[col] });
        const cell = sheet[cellAddr];
        const v = cell && cell.v != null ? parseFloat(cell.v) : null;
        if (v !== null && !isNaN(v)) {
            sums[col] += v;
            counts[col]++;
        }
    }
}

const avgs = {};
for (const col of ['宝贝质量', '物流速度', '服务保障']) {
    avgs[col] = Math.round((sums[col] / counts[col]) * 10) / 10;
    console.log(col + ': count=' + counts[col] + ', avg=' + avgs[col]);
}

// 填充缺失值（行23和行72，即Excel第24和73行）
const missingRows = [23, 72];
for (const row of missingRows) {
    for (const col of ['宝贝质量', '物流速度', '服务保障']) {
        const cellAddr = XLSX.utils.encode_cell({ r: row, c: colIdx[col] });
        const cell = sheet[cellAddr];
        const currentVal = cell && cell.v != null ? parseFloat(cell.v) : null;
        if (currentVal === null || isNaN(currentVal)) {
            if (!sheet[cellAddr]) sheet[cellAddr] = {};
            sheet[cellAddr].v = avgs[col];
            const idCellAddr = XLSX.utils.encode_cell({ r: row, c: colIdx['商品id'] });
            console.log('填充 ID=' + sheet[idCellAddr].v + ' 的 ' + col + ' = ' + avgs[col]);
        }
    }
}

XLSX.writeFile(wb, path);
console.log('\n已完成填充并保存！');
