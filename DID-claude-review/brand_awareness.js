const fs = require('fs');
const XLSX = require('xlsx');

// 1. 收集百度指数CSV的品牌->指数映射
const csvPath = 'C:\\DEV\\DID\\京东商品百度指数.csv';
const content = fs.readFileSync(csvPath, 'utf8');
const normalized = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
const simplified = normalized.replace(/"[^\"]*"/g, (m) => m.replace(/\s/g, ''));
const lines = simplified.split('\n').filter(l => l.trim());

const baiduMap = {};
for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line || line === ',') continue;
    const commaIdx = line.indexOf(',');
    if (commaIdx === -1) continue;
    const original = line.slice(0, commaIdx).trim();
    const valueRaw = line.slice(commaIdx + 1).replace(/"/g, '').trim();
    if (!original) continue;
    const clean = original.replace(/-化妆品$/, '');
    const value = parseFloat(valueRaw.replace(/,/g, '')) || null;
    if (value !== null) baiduMap[clean] = { original, value };
}
console.log('百度指数CSV品牌数:', Object.keys(baiduMap).length);

// 2. 读取商品详情
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
console.log('品牌列:', brandIdx, '百度指数列:', baiduIdx);

// 3. 将9个用平均值填充的改回null，收集所有有效值
const avgValue = 102894;
const validValues = new Set(); // 用Set去重

for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const brandCellAddr = XLSX.utils.encode_cell({ r: row, c: brandIdx });
    const indexCellAddr = XLSX.utils.encode_cell({ r: row, c: baiduIdx });
    const brandCell = sheet[brandCellAddr];
    const indexCell = sheet[indexCellAddr];
    const brand = brandCell && brandCell.v != null ? String(brandCell.v).trim() : '';
    const indexVal = indexCell && indexCell.v != null ? parseFloat(indexCell.v) : null;

    if (indexVal === avgValue) {
        // 用平均值填充的，改回null
        sheet[indexCellAddr] = { v: null };
    } else if (indexVal !== null && !isNaN(indexVal)) {
        validValues.add(indexVal);
    }
}

const sortedUniqueValues = [...validValues].sort((a, b) => b - a);
const topThirdCount = Math.ceil(sortedUniqueValues.length / 3);
const thresholdValue = sortedUniqueValues[topThirdCount - 1];

console.log('\n去重后有效指数值数:', sortedUniqueValues.length);
console.log('Top 1/3数量:', topThirdCount);
console.log('阈值:', thresholdValue);
console.log('\n所有去重值（降序）:', sortedUniqueValues);

// 4. 收集每个有效值对应的品牌（用于展示）
const valueToBrands = {};
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const brandCellAddr = XLSX.utils.encode_cell({ r: row, c: brandIdx });
    const indexCellAddr = XLSX.utils.encode_cell({ r: row, c: baiduIdx });
    const brandCell = sheet[brandCellAddr];
    const indexCell = sheet[indexCellAddr];
    const brand = brandCell && brandCell.v != null ? String(brandCell.v).trim() : '';
    const indexVal = indexCell && indexCell.v != null ? parseFloat(indexCell.v) : null;

    if (indexVal !== null && !isNaN(indexVal) && indexVal !== avgValue) {
        if (!valueToBrands[indexVal]) valueToBrands[indexVal] = [];
        valueToBrands[indexVal].push(brand);
    }
}

console.log('\nTop 1/3品牌（brand_awareness=1）:');
let awareness1Count = 0;
for (const v of sortedUniqueValues) {
    if (v >= thresholdValue) {
        awareness1Count += valueToBrands[v].length;
        console.log('  ' + v + ': ' + valueToBrands[v].join(', '));
    }
}
console.log('brand_awareness=1 的行数:', awareness1Count);

// 5. 添加brand_awareness列并填充
const newColIdx = range.e.c + 1;
sheet[XLSX.utils.encode_cell({ r: 0, c: newColIdx })] = { v: 'brand_awareness' };

let cnt1 = 0, cnt0 = 0, cntNull = 0;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const indexCellAddr = XLSX.utils.encode_cell({ r: row, c: baiduIdx });
    const indexCell = sheet[indexCellAddr];
    const indexVal = indexCell && indexCell.v != null ? parseFloat(indexCell.v) : null;
    const newCellAddr = XLSX.utils.encode_cell({ r: row, c: newColIdx });

    if (indexVal === null || isNaN(indexVal)) {
        sheet[newCellAddr] = { v: null };
        cntNull++;
    } else if (indexVal >= thresholdValue) {
        sheet[newCellAddr] = { v: 1 };
        cnt1++;
    } else {
        sheet[newCellAddr] = { v: 0 };
        cnt0++;
    }
}

range.e.c = newColIdx;
sheet['!ref'] = XLSX.utils.encode_range(range);
XLSX.writeFile(wb, './商品详情_清洗后.xlsx');

console.log('\n结果: 1=' + cnt1 + ', 0=' + cnt0 + ', null=' + cntNull);
console.log('完成！');
