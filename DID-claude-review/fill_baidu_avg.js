const fs = require('fs');
const XLSX = require('xlsx');

// 1. 收集商品详情_清洗后.xlsx 中所有非空百度指数值
const wb = XLSX.readFile('./商品详情_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : 'col_' + col);
}
const baiduIdx = headers.indexOf('百度指数');
console.log('百度指数列索引:', baiduIdx);

const detailValues = [];
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const c = sheet[XLSX.utils.encode_cell({ r: row, c: baiduIdx })];
    const v = c && c.v != null ? parseFloat(c.v) : null;
    if (v !== null && !isNaN(v)) detailValues.push(v);
}
console.log('详情表有效百度指数数:', detailValues.length);

// 2. 收集百度指数CSV中所有值
const path = 'C:\\DEV\\DID\\京东商品百度指数.csv';
const content = fs.readFileSync(path, 'utf8');
const normalized = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
const simplified = normalized.replace(/"[^\"]*"/g, (m) => m.replace(/\s/g, ''));
const lines = simplified.split('\n').filter(l => l.trim());

const csvValues = [];
for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line || line === ',') continue;
    const commaIdx = line.indexOf(',');
    if (commaIdx === -1) continue;
    const valueRaw = line.slice(commaIdx + 1).replace(/"/g, '').trim();
    const value = parseFloat(valueRaw.replace(/,/g, '')) || null;
    if (value !== null && !isNaN(value)) csvValues.push(value);
}
console.log('百度指数CSV有效值数:', csvValues.length);

// 3. 合并去重计算平均值
const allValues = [...detailValues, ...csvValues];
const uniqueValues = [...new Set(allValues)];
const sum = uniqueValues.reduce((a, b) => a + b, 0);
const avg = Math.round(sum / uniqueValues.length);
console.log('\n合并去重后总有效值数:', uniqueValues.length);
console.log('平均值:', avg);

// 4. 填充9个空值行
const brandIdx = headers.indexOf('品牌');
const emptyBrandRows = [];
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const brandCell = sheet[XLSX.utils.encode_cell({ r: row, c: brandIdx })];
    const indexCell = sheet[XLSX.utils.encode_cell({ r: row, c: baiduIdx })];
    const brand = brandCell && brandCell.v != null ? String(brandCell.v) : '';
    const indexVal = indexCell && indexCell.v != null ? parseFloat(indexCell.v) : null;
    if (indexVal === null || isNaN(indexVal)) {
        emptyBrandRows.push({ row: row + 1, brand });
        const cellAddr = XLSX.utils.encode_cell({ r: row, c: baiduIdx });
        sheet[cellAddr] = { v: avg };
        console.log('填充 Row' + (row+1) + ' (' + brand + ') -> ' + avg);
    }
}

XLSX.writeFile(wb, './商品详情_清洗后.xlsx');
console.log('\n填充完成，已保存！共填充', emptyBrandRows.length, '行');
