const XLSX = require('xlsx');

const wb = XLSX.readFile('./商品详情_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

// 删除重复的 brand_awareness 列（保留最后一个，即col_12，删除col_11）
// 先找到两列的索引
const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : 'col_' + col);
}

const indices = [];
for (let i = 0; i < headers.length; i++) {
    if (headers[i] === 'brand_awareness') indices.push(i);
}
console.log('brand_awareness出现的列索引:', indices);

// 删除第一个（重复的旧列，索引11），保留最后一个（索引12）
const toDelete = indices[0];
console.log('删除列索引:', toDelete, '(Excel列' + (toDelete+1) + ')');

// 删除该列：将该列之后的所有列数据左移
for (let row = range.s.r; row <= range.e.r; row++) {
    for (let col = toDelete; col < range.e.c; col++) {
        const srcAddr = XLSX.utils.encode_cell({ r: row, c: col + 1 });
        const dstAddr = XLSX.utils.encode_cell({ r: row, c: col });
        sheet[dstAddr] = sheet[srcAddr];
    }
    // 清空最后一列
    sheet[XLSX.utils.encode_cell({ r: row, c: range.e.c })] = { v: null };
}

range.e.c -= 1;
sheet['!ref'] = XLSX.utils.encode_range(range);

// 验证
const newHeaders = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    newHeaders.push(c && c.v != null ? String(c.v) : 'col_' + col);
}
console.log('\n删除重复列后的表头:', newHeaders);

// 检查brand_awareness列的值分布
const awarenessIdx = newHeaders.indexOf('brand_awareness');
console.log('brand_awareness列索引:', awarenessIdx);
const stats = { '1': 0, '0': 0, 'null': 0};
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const a = sheet[XLSX.utils.encode_cell({ r: row, c: awarenessIdx })];
    const v = a && a.v != null ? String(a.v) : 'null';
    if (v === '1') stats['1']++;
    else if (v === '0') stats['0']++;
    else stats['null']++;
}
console.log('统计: 1=' + stats['1'] + ' 0=' + stats['0'] + ' null=' + stats['null']);

// 检查百度指数为null的行
const baiduIdx = newHeaders.indexOf('百度指数');
console.log('\n百度指数为null的行:');
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const bi = sheet[XLSX.utils.encode_cell({ r: row, c: baiduIdx })];
    const ai = sheet[XLSX.utils.encode_cell({ r: row, c: awarenessIdx })];
    const br = sheet[XLSX.utils.encode_cell({ r: row, c: newHeaders.indexOf('品牌') })];
    if (bi && (bi.v === null || bi.v === '')) {
        console.log('  Row' + (row+1) + ' brand=' + (br && br.v ? br.v : '') + ' awareness=' + (ai && ai.v !== null ? ai.v : 'null'));
    }
}

XLSX.writeFile(wb, './商品详情_清洗后.xlsx');
console.log('\n完成，已保存！');
