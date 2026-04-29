const XLSX = require('xlsx');
const fs = require('fs');

const detailPath = './商品详情.xlsx';
const commentPath = './评论数.xlsx';
const outPath = './商品详情_合并后.xlsx';

const wbDetail = XLSX.readFile(detailPath, { cellNF: true, cellDates: true });
const wbComment = XLSX.readFile(commentPath, { cellNF: true, cellDates: true });

// ---- 读取详情表 ----
const sheetD = wbDetail.Sheets[wbDetail.SheetNames[0]];
const rangeD = XLSX.utils.decode_range(sheetD['!ref']);

// 获取表头
const headers = [];
for (let col = rangeD.s.c; col <= rangeD.e.c; col++) {
    const cell = sheetD[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(cell && cell.v != null ? String(cell.v) : `col_${col}`);
}

// 读取所有数据行
const detailRows = []; // { id, rowIndex, values: [] }
for (let row = rangeD.s.r + 1; row <= rangeD.e.r; row++) {
    const idCell = sheetD[XLSX.utils.encode_cell({ r: row, c: 0 })];
    const id = idCell && idCell.v != null ? String(idCell.v) : null;
    const values = [];
    for (let col = rangeD.s.c; col <= rangeD.e.c; col++) {
        const c = sheetD[XLSX.utils.encode_cell({ r: row, c: col })];
        values.push(c && c.v != null ? c.v : null);
    }
    detailRows.push({ id, rowIndex: row, values });
}

// 按ID分组，找出价格列（列索引4，即"价格"列）
const priceColIndex = 4; // 价格在第5列（从0开始）

const grouped = {};
for (const row of detailRows) {
    if (!grouped[row.id]) grouped[row.id] = [];
    grouped[row.id].push(row);
}

// 对每个ID，按价格取中位数行
const medianRows = [];
for (const [id, rows] of Object.entries(grouped)) {
    if (rows.length === 1) {
        medianRows.push(rows[0]);
    } else {
        // 提取价格并排序，取中位数
        const sorted = rows.slice().sort((a, b) => {
            const pa = a.values[priceColIndex] ?? -Infinity;
            const pb = b.values[priceColIndex] ?? -Infinity;
            return pa - pb;
        });
        const midIdx = Math.floor(sorted.length / 2);
        medianRows.push(sorted[midIdx]);
    }
}

console.log(`详情表原始行数: ${detailRows.length}`);
console.log(`groupby后行数: ${medianRows.length}`);

// 构建新工作表（详情表 + 评论数列，但去掉重复的ID列）
const newHeaders = [...headers, '评论数'];
const newData = [newHeaders];

// 读取评论表（第一列是ID，第二列是评论数）
const sheetC = wbComment.Sheets[wbComment.SheetNames[0]];
const rangeC = XLSX.utils.decode_range(sheetC['!ref']);
const commentMap = {};
for (let row = rangeC.s.r + 1; row <= rangeC.e.r; row++) {
    const idCell = sheetC[XLSX.utils.encode_cell({ r: row, c: 0 })];
    const countCell = sheetC[XLSX.utils.encode_cell({ r: row, c: 1 })];
    const id = idCell && idCell.v != null ? String(idCell.v) : null;
    const count = countCell && countCell.v != null ? countCell.v : null;
    if (id) commentMap[id] = count;
}

// 按ID排序输出（与评论表顺序一致）
const sortedMedianRows = medianRows.slice().sort((a, b) => String(a.id).localeCompare(String(b.id)));
let addedCount = 0;
for (const row of sortedMedianRows) {
    const commentCount = commentMap[row.id] ?? null;
    newData.push([...row.values, commentCount]);
    addedCount++;
}

// 写入新文件
const newWb = XLSX.utils.book_new();
const newWs = XLSX.utils.aoa_to_sheet(newData);
XLSX.utils.book_append_sheet(newWb, newWs, '合并结果');
XLSX.writeFile(newWb, outPath);

console.log(`合并后行数（含表头）: ${newData.length}`);
console.log(`输出文件: ${outPath}`);
console.log('完成！');
