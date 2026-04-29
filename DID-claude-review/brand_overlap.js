const fs = require('fs');
const XLSX = require('xlsx');

// 读取商品详情_清洗后.xlsx 的品牌列
const wb = XLSX.readFile('./商品详情_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : 'col_' + col);
}
const brandIdx = headers.indexOf('品牌');
console.log('品牌列索引:', brandIdx);

const detailBrands = new Set();
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const c = sheet[XLSX.utils.encode_cell({ r: row, c: brandIdx })];
    const v = c && c.v != null ? String(c.v).trim() : '';
    if (v) detailBrands.add(v);
}
console.log('商品详情品牌数:', detailBrands.size);
console.log('品牌列表:', [...detailBrands]);

// 读取百度指数CSV
const path = 'C:\\DEV\\DID\\京东商品百度指数.csv';
const content = fs.readFileSync(path, 'utf8');
const normalized = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
const simplified = normalized.replace(/"[^\"]*"/g, (m) => m.replace(/\s/g, ''));
const lines = simplified.split('\n').filter(l => l.trim());

const baiduBrands = new Set();
for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line || line === ',') continue;
    const commaIdx = line.indexOf(',');
    if (commaIdx === -1) continue;
    const keyword = line.slice(0, commaIdx).trim();
    if (keyword) baiduBrands.add(keyword);
}
console.log('\n百度指数品牌数:', baiduBrands.size);
console.log('品牌列表:', [...baiduBrands]);

// 匹配：商品详情品牌是否包含百度指数的关键词（去掉"-化妆品"后缀）
const matched = [];
const unmatched = [];
for (const b of baiduBrands) {
    const cleanB = b.replace(/-化妆品$/, '');
    const found = [...detailBrands].some(d => d.includes(cleanB) || cleanB.includes(d));
    if (found) matched.push(b);
    else unmatched.push(b);
}

console.log('\n=== 重叠分析 ===');
console.log('百度指数品牌与商品详情品牌匹配数:', matched.length, '/', baiduBrands.size);
console.log('\n已匹配:', matched);
console.log('\n未匹配:', unmatched);
