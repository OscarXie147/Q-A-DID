const XLSX = require('xlsx');

const qaPath = './问答_符合条件.xlsx';
const detailPath = './商品详情_清洗后.xlsx';
const outPath = './问答_合并商品信息.xlsx';

// 读取问答表
const wbQa = XLSX.readFile(qaPath, { cellNF: true, cellDates: true });
const sheetQa = wbQa.Sheets[wbQa.SheetNames[0]];
const rangeQa = XLSX.utils.decode_range(sheetQa['!ref']);

const qaHeaders = [];
for (let col = rangeQa.s.c; col <= rangeQa.e.c; col++) {
    const c = sheetQa[XLSX.utils.encode_cell({ r: 0, c: col })];
    qaHeaders.push(c && c.v != null ? String(c.v).trim() : 'col_' + col);
}
console.log('问答表表头:', qaHeaders);

// 读取商品详情表
const wbDetail = XLSX.readFile(detailPath, { cellNF: true, cellDates: true });
const sheetDetail = wbDetail.Sheets[wbDetail.SheetNames[0]];
const rangeDetail = XLSX.utils.decode_range(sheetDetail['!ref']);

const detailHeaders = [];
for (let col = rangeDetail.s.c; col <= rangeDetail.e.c; col++) {
    const c = sheetDetail[XLSX.utils.encode_cell({ r: 0, c: col })];
    detailHeaders.push(c && c.v != null ? String(c.v).trim() : 'col_' + col);
}
console.log('商品详情表表头:', detailHeaders);

// 商品详情的ID列索引
const detailIdIdx = detailHeaders.indexOf('商品id');
console.log('商品详情ID列索引:', detailIdIdx);

// 构建商品详情 lookup: id -> {col -> value}，去掉ID列
const detailLookup = {};
for (let row = rangeDetail.s.r + 1; row <= rangeDetail.e.r; row++) {
    const idCell = sheetDetail[XLSX.utils.encode_cell({ r: row, c: detailIdIdx })];
    const id = idCell && idCell.v != null ? String(idCell.v) : null;
    if (!id) continue;
    if (!detailLookup[id]) detailLookup[id] = {};
    for (let col = rangeDetail.s.c; col <= rangeDetail.e.c; col++) {
        if (col === detailIdIdx) continue; // 跳过ID列
        const h = detailHeaders[col];
        const c = sheetDetail[XLSX.utils.encode_cell({ r: row, c: col })];
        detailLookup[id][h] = c && c.v !== undefined && c.v !== null ? c.v : null;
    }
}
console.log('商品详情lookup行数:', Object.keys(detailLookup).length);

// 新表头：问答表全部列 + 商品详情非ID列
const newHeaders = [...qaHeaders, ...detailHeaders.filter((_, i) => i !== detailIdIdx)];
console.log('新表头列数:', newHeaders.length);

// 遍历问答表，每行按ID匹配商品详情
const newData = [newHeaders];
let matched = 0, unmatched = 0;

for (let row = rangeQa.s.r + 1; row <= rangeQa.e.r; row++) {
    const idCell = sheetQa[XLSX.utils.encode_cell({ r: row, c: 0 })];
    const id = idCell && idCell.v != null ? String(idCell.v) : null;

    // 问答表该行数据
    const qaRow = [];
    for (let col = rangeQa.s.c; col <= rangeQa.e.c; col++) {
        const c = sheetQa[XLSX.utils.encode_cell({ r: row, c: col })];
        qaRow.push(c && c.v !== undefined && c.v !== null ? c.v : null);
    }

    // 商品详情匹配数据
    if (id && detailLookup[id]) {
        matched++;
        const detailRow = detailHeaders
            .map((h, i) => ({ h, i }))
            .filter(({ i }) => i !== detailIdIdx)
            .map(({ h }) => detailLookup[id][h] ?? null);
        newData.push([...qaRow, ...detailRow]);
    } else {
        unmatched++;
        const detailRow = detailHeaders
            .map((h, i) => ({ h, i }))
            .filter(({ i }) => i !== detailIdIdx)
            .map(() => null);
        newData.push([...qaRow, ...detailRow]);
    }
}

console.log('\n匹配成功:', matched, '行');
console.log('未匹配:', unmatched, '行');

const newWb = XLSX.utils.book_new();
const newWs = XLSX.utils.aoa_to_sheet(newData);
XLSX.utils.book_append_sheet(newWb, newWs, '问答_合并商品信息');
XLSX.writeFile(newWb, outPath);
console.log('\n输出文件:', outPath);
console.log('总行数（含表头）:', newData.length);
console.log('总列数:', newHeaders.length);
console.log('新表头:', newHeaders.join(' | '));
