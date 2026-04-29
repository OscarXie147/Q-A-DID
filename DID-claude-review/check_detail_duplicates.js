const XLSX = require('xlsx');
const fs = require('fs');

const detailPath = './商品详情.xlsx';

const wb = XLSX.readFile(detailPath, { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

// 收集每行第一列的ID
const idCount = {};
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
    if (cell && cell.v !== null && cell.v !== undefined) {
        const id = cell.v;
        idCount[id] = (idCount[id] || 0) + 1;
    }
}

const ids = Object.keys(idCount);
const counts = Object.values(idCount);
const totalIds = ids.length;
const totalRows = counts.reduce((a, b) => a + b, 0);
const minCount = Math.min(...counts);
const maxCount = Math.max(...counts);
const avgCount = (totalRows / totalIds).toFixed(2);

let result = '';
result += `商品详情.xlsx 总行数: ${totalRows}\n`;
result += `唯一ID数: ${totalIds}\n`;
result += `每个ID的行数: min=${minCount}, max=${maxCount}, avg=${avgCount}\n\n`;

// 按行数分组统计
const groupByCount = {};
for (const [id, cnt] of Object.entries(idCount)) {
    groupByCount[cnt] = (groupByCount[cnt] || 0) + 1;
}
result += '每个ID的行数分布:\n';
for (const [cnt, num] of Object.entries(groupByCount).sort((a, b) => a[0] - b[0])) {
    result += `  ${cnt}行的ID数: ${num}\n`;
}

// 显示每个ID的行数
result += '\n每个ID的行数:\n';
for (const [id, cnt] of Object.entries(idCount).sort((a, b) => a[0] - b[0])) {
    result += `  ID ${id}: ${cnt}行\n`;
}

fs.writeFileSync('./detail_row_per_id.txt', result, 'utf-8');
console.log(result);
