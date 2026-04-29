const fs = require('fs');
const XLSX = require('xlsx');

// 读取百度指数品牌
const path = 'C:\\DEV\\DID\\京东商品百度指数.csv';
const content = fs.readFileSync(path, 'utf8');
const normalized = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
const simplified = normalized.replace(/"[^\"]*"/g, (m) => m.replace(/\s/g, ''));
const lines = simplified.split('\n').filter(l => l.trim());

const baiduBrands = []; // { original: 'xxx-化妆品', clean: 'xxx' }
for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line || line === ',') continue;
    const commaIdx = line.indexOf(',');
    if (commaIdx === -1) continue;
    const original = line.slice(0, commaIdx).trim();
    if (!original) continue;
    const clean = original.replace(/-化妆品$/, '');
    baiduBrands.push({ original, clean });
}
console.log('百度指数品牌数:', baiduBrands.length);

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
console.log('品牌列索引:', brandIdx);

// 对每个百度品牌，提取可匹配的关键词
// 比如 "欧莱雅-化妆品" -> ["欧莱雅", "l'oreal"] -> 用于子串匹配
function getMatchKeys(brand) {
    const keys = [];
    keys.push(brand.clean.toLowerCase());
    // 如果是 "xxx/中文" 格式，提取英文部分
    if (brand.clean.includes('/')) {
        const parts = brand.clean.split('/');
        keys.push(parts[0].toLowerCase());
        keys.push(parts[1] ? parts[1].toLowerCase() : '');
    }
    return keys;
}

const baiduMatchEntries = baiduBrands.map(b => ({
    ...b,
    keys: getMatchKeys(b)
}));

// 检查商品详情中的每个品牌
const notFound = [];
const brandChanges = [];

for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const cellAddr = XLSX.utils.encode_cell({ r: row, c: brandIdx });
    const cell = sheet[cellAddr];
    const brandRaw = cell && cell.v != null ? String(cell.v).trim() : '';

    let found = false;
    for (const b of baiduMatchEntries) {
        const detailLower = brandRaw.toLowerCase();
        // 检查百度品牌的clean名是否包含在详情品牌中，或者反过来
        const cleanLower = b.clean.toLowerCase();
        const keysLower = b.keys.map(k => k.toLowerCase());

        const matches = keysLower.some(k => k && (detailLower.includes(k) || k.includes(detailLower)));

        if (matches) {
            found = true;
            // 如果详情品牌和百度指数的写法不一样，更新
            if (brandRaw !== b.original) {
                brandChanges.push({ row: row + 1, old: brandRaw, new: b.original });
                sheet[cellAddr].v = b.original;
            }
            break;
        }
    }

    if (!found) {
        notFound.push({ row: row + 1, brand: brandRaw });
    }
}

console.log('\n=== 结果 ===');
console.log('商品详情品牌总数: 99');
console.log('在百度指数中找到的: ' + (99 - notFound.length));
console.log('在百度指数中找不到的: ' + notFound.length);

if (notFound.length > 0) {
    console.log('\n找不到的brand:');
    for (const n of notFound) {
        console.log('  Row' + n.row + ': ' + n.brand);
    }
}

if (brandChanges.length > 0) {
    console.log('\n品牌写法已更新:');
    for (const c of brandChanges) {
        console.log('  Row' + c.row + ': ' + c.old + ' -> ' + c.new);
    }
}

XLSX.writeFile(wb, './商品详情_清洗后.xlsx');
console.log('\n文件已保存！');
