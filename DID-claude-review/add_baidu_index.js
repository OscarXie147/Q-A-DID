const fs = require('fs');
const XLSX = require('xlsx');

// 读取百度指数，建立 clean名 -> 指数 的映射
const path = 'C:\\DEV\\DID\\京东商品百度指数.csv';
const content = fs.readFileSync(path, 'utf8');
const normalized = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
const simplified = normalized.replace(/"[^\"]*"/g, (m) => m.replace(/\s/g, ''));
const lines = simplified.split('\n').filter(l => l.trim());

const baiduMap = {}; // clean名 -> 指数值
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
    baiduMap[clean] = value;
}

console.log('百度指数映射数:', Object.keys(baiduMap).length);

// 读取商品详情
const wb = XLSX.readFile('./商品详情_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : 'col_' + col);
}
const brandIdx = headers.indexOf('品牌');
console.log('品牌列索引:', brandIdx, '当前列数:', headers.length);

// 在最后追加"百度指数"列
const newColIdx = range.e.c + 1;
const newHeaderAddr = XLSX.utils.encode_cell({ r: 0, c: newColIdx });
sheet[newHeaderAddr] = { v: '百度指数' };

// 建立品牌名 -> 百度指数clean名 的匹配表（复用之前的逻辑）
function getMatchKeys(brandClean) {
    const keys = [];
    keys.push(brandClean.toLowerCase());
    if (brandClean.includes('/')) {
        const parts = brandClean.split('/');
        keys.push(parts[0].toLowerCase());
        keys.push(parts[1] ? parts[1].toLowerCase() : '');
    }
    return keys;
}

// 预处理百度Map的keys
const baiduEntries = Object.keys(baiduMap).map(k => ({
    clean: k,
    keys: getMatchKeys(k),
    value: baiduMap[k]
}));

// 遍历每行，填充百度指数
let filled = 0;
let empty = 0;
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const cellAddr = XLSX.utils.encode_cell({ r: row, c: brandIdx });
    const cell = sheet[cellAddr];
    const brandRaw = cell && cell.v != null ? String(cell.v).trim() : '';
    const brandClean = brandRaw.replace(/-化妆品$/, '').replace(/[\[\]']/g, '').trim();
    const brandLower = brandClean.toLowerCase();

    let foundValue = null;
    for (const b of baiduEntries) {
        const matches = b.keys.some(k => k && (brandLower.includes(k) || k.includes(brandLower)));
        if (matches) {
            foundValue = b.value;
            break;
        }
    }

    const newCellAddr = XLSX.utils.encode_cell({ r: row, c: newColIdx });
    if (foundValue !== null) {
        sheet[newCellAddr] = { v: foundValue };
        filled++;
    } else {
        sheet[newCellAddr] = { v: null };
        empty++;
    }
}

// 更新range的末尾列
range.e.c = newColIdx;
sheet['!ref'] = XLSX.utils.encode_range(range);

XLSX.writeFile(wb, './商品详情_清洗后.xlsx');
console.log('\n完成！');
console.log('已填充百度指数:', filled);
console.log('空值（未匹配）:', empty);
console.log('输出文件: ./商品详情_清洗后.xlsx');

// 验证打印前5行
console.log('\n前5行品牌+百度指数:');
for (let row = range.s.r + 1; row <= range.s.r + 5; row++) {
    const brandCell = sheet[XLSX.utils.encode_cell({ r: row, c: brandIdx })];
    const indexCell = sheet[XLSX.utils.encode_cell({ r: row, c: newColIdx })];
    console.log('  ' + (brandCell && brandCell.v ? brandCell.v : '') + ' -> ' + (indexCell && indexCell.v !== null ? indexCell.v : '(空)'));
}
