const XLSX = require('xlsx');
const fs = require('fs');

const detailPath = './商品详情.xlsx';
const wb = XLSX.readFile(detailPath, { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

// 读取表头（第一行）
const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(cell && cell.v !== null && cell.v !== undefined ? String(cell.v) : `col_${col}`);
}

let result = '';
result += `总列数: ${headers.length}\n\n`;

// 对每一列，统计：唯一值数量、空值数量、每个ID是否唯一（group by第一列ID时是否一致）
const idCell = sheet[XLSX.utils.encode_cell({ r: 0, c: 0 })];
result += `第一列ID标题: ${headers[0]}\n\n`;

// 收集每列数据
const colData = {};
for (let col = 0; col < headers.length; col++) {
    colData[col] = { values: new Set(), nullCount: 0, valueCounts: {} };
}

const idToColValues = {};

for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const idCellVal = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
    const id = idCellVal ? idCellVal.v : null;

    for (let col = 0; col < headers.length; col++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
        const val = cell && cell.v !== null && cell.v !== undefined ? String(cell.v).trim() : '__NULL__';
        colData[col].values.add(val);
        if (val === '__NULL__') colData[col].nullCount++;
        const key = `${col}_${val}`;
        colData[col].valueCounts[key] = (colData[col].valueCounts[key] || 0) + 1;
    }
}

// 按ID分组，看每列在同ID内是否一致
const idGroups = {};
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
    const id = cell && cell.v !== null ? String(cell.v) : 'unknown';
    if (!idGroups[id]) idGroups[id] = {};
    for (let col = 0; col < headers.length; col++) {
        const c = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
        const val = c && c.v !== null && c.v !== undefined ? String(c.v).trim() : '__NULL__';
        if (!idGroups[id][col]) idGroups[id][col] = new Set();
        idGroups[id][col].add(val);
    }
}

// 对于每列，检查group by ID后是否唯一
result += '各列详情（group by ID后是否唯一、空值数、唯一值数）:\n';
const colStats = [];
for (let col = 0; col < headers.length; col++) {
    let allSame = true;
    let idCount = 0;
    for (const [id, cols] of Object.entries(idGroups)) {
        idCount++;
        if (cols[col] && cols[col].size > 1) {
            allSame = false;
            break;
        }
    }
    const uniqueCount = colData[col].values.size;
    const nullCount = colData[col].nullCount;
    const isUniqueWhenGrouped = allSame ? '是' : '否';
    colStats.push({
        col,
        header: headers[col],
        isUniqueWhenGrouped,
        nullCount,
        uniqueCount,
        allSame
    });
}

// 打印
for (const s of colStats) {
    result += `\n列${s.col}: ${s.header}\n`;
    result += `  groupby ID后唯一: ${s.isUniqueWhenGrouped}\n`;
    result += `  空值行数: ${s.nullCount}\n`;
    result += `  唯一值数: ${s.uniqueCount}\n`;
    if (!s.allSame) {
        result += `  注意: 同一ID下有多个不同值！\n`;
    }
}

// 列出不唯一的列的详细情况
result += '\n\n--- groupby ID后不唯一的列详情 ---\n';
for (const s of colStats) {
    if (!s.allSame) {
        result += `\n列${s.col}: ${s.header}\n`;
        for (const [id, cols] of Object.entries(idGroups)) {
            if (cols[s.col] && cols[s.col].size > 1) {
                result += `  ID ${id}: ${[...cols[s.col]].join(' | ')}\n`;
            }
        }
    }
}

fs.writeFileSync('./detail_columns_analysis.txt', result, 'utf-8');
console.log(result);
