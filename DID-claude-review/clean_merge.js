const XLSX = require('xlsx');
const fs = require('fs');

const inPath = './商品详情_合并后.xlsx';
const outPath = './商品详情_清洗后.xlsx';

const wb = XLSX.readFile(inPath, { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

// 读取所有数据
const rawHeaders = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    rawHeaders.push(c && c.v != null ? String(c.v) : `col_${col}`);
}

// 列索引
const idx = {
    商品id: rawHeaders.indexOf('商品id'),
    标题: rawHeaders.indexOf('标题'),
    skuid: rawHeaders.indexOf('skuid'),
    sku: rawHeaders.indexOf('sku'),
    价格: rawHeaders.indexOf('价格'),
    券后价格: rawHeaders.indexOf('券后价格'),
    店铺: rawHeaders.indexOf('店铺'),
    店铺评分: rawHeaders.indexOf('店铺评分'),
    服务: rawHeaders.indexOf('服务'),
    品牌: rawHeaders.indexOf('品牌'),
    评论数: rawHeaders.indexOf('评论数'),
};

console.log('列索引:', idx);

// 解析评论数字符串，如"1万+" -> 1, "6000+" -> 0.6
function parseCommentCount(s) {
    if (s == null || s === '(空)') return null;
    const str = String(s).trim();
    if (str.includes('万')) {
        return parseFloat(str.replace('万+', '')) || null;
    } else {
        const n = parseFloat(str.replace(/[+人]/g, ''));
        return n ? n / 10000 : null;
    }
}

// 解析店铺评分JSON数组（可能用单引号）
function parseShopScore(s) {
    if (s == null || s === '(空)') return { 宝贝质量: null, 物流速度: null, 服务保障: null };
    let str = String(s);
    try {
        // 单引号JSON，转成双引号再解析
        str = str.replace(/'/g, '"');
        str = JSON.parse(str);
    } catch {}
    const result = { 宝贝质量: null, 物流速度: null, 服务保障: null };
    if (Array.isArray(str)) {
        for (const item of str) {
            if (item.title === '宝贝质量') result.宝贝质量 = parseFloat(item.score);
            if (item.title === '物流速度') result.物流速度 = parseFloat(item.score);
            if (item.title === '服务保障') result.服务保障 = parseFloat(item.score);
        }
    }
    return result;
}

// 清洗服务/品牌列：去掉 [] ''，保留内容用逗号分隔
function cleanArrayStr(s) {
    if (s == null || s === '(空)') return '';
    let str = String(s).trim();
    // 如果是JSON数组格式（单引号），转双引号后解析
    try {
        str = str.replace(/'/g, '"');
        str = JSON.parse(str);
    } catch {}
    if (Array.isArray(str)) {
        return str.join(',');
    }
    // 否则去掉 ' 和 [ ]
    str = str.replace(/[\[\]']/g, '');
    return str;
}

// 读取所有数据行
const dataRows = [];
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const get = (colIdx) => {
        const c = sheet[XLSX.utils.encode_cell({ r: row, c: colIdx })];
        return c && c.v != null ? c.v : null;
    };

    const 商品id = get(idx.商品id);
    const 标题 = get(idx.标题);
    const 价格 = get(idx.价格);
    const 店铺 = get(idx.店铺);
    const 评论数原始 = get(idx.评论数);
    const 服务原始 = get(idx.服务);
    const 品牌原始 = get(idx.品牌);

    const 评论数_万 = parseCommentCount(评论数原始);
    const scoreParsed = parseShopScore(get(idx.店铺评分));
    const 服务 = cleanArrayStr(服务原始);
    const 品牌 = cleanArrayStr(品牌原始);

    dataRows.push({
        商品id, 标题, 价格, 店铺,
        宝贝质量: scoreParsed.宝贝质量,
        物流速度: scoreParsed.物流速度,
        服务保障: scoreParsed.服务保障,
        服务,
        品牌,
        '评论数（万）': 评论数_万,
    });
}

// 新表头
const newHeaders = ['商品id', '标题', '价格', '店铺', '宝贝质量', '物流速度', '服务保障', '服务', '品牌', '评论数（万）'];

const newData = [newHeaders];
for (const row of dataRows) {
    newData.push([
        row.商品id,
        row.标题,
        row.价格,
        row.店铺,
        row.宝贝质量,
        row.物流速度,
        row.服务保障,
        row.服务,
        row.品牌,
        row['评论数（万）'],
    ]);
}

const newWb = XLSX.utils.book_new();
const newWs = XLSX.utils.aoa_to_sheet(newData);
XLSX.utils.book_append_sheet(newWb, newWs, '清洗后');
XLSX.writeFile(newWb, outPath);

console.log(`输出行数（含表头）: ${newData.length}`);
console.log(`新表头: ${newHeaders.join(' | ')}`);
console.log('\n前5行预览:');
for (let i = 1; i <= Math.min(5, newData.length - 1); i++) {
    console.log(`Row${i}: ${newData[i].join(' | ')}`);
}

console.log('\n完成！输出文件:', outPath);
