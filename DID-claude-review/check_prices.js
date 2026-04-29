const XLSX = require('xlsx');

const wb = XLSX.readFile('./问答_合并商品信息.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v !== undefined ? String(c.v).trim() : 'col_' + col);
}
const priceIdx = headers.indexOf('价格');
const idIdx = 0;

console.log('价格列索引:', priceIdx);

// 收集每个商品的ID和价格
const productPrices = {};
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const idCell = sheet[XLSX.utils.encode_cell({ r: row, c: idIdx })];
    const priceCell = sheet[XLSX.utils.encode_cell({ r: row, c: priceIdx })];
    const id = idCell && idCell.v !== undefined ? String(idCell.v) : null;
    const v = priceCell && priceCell.v !== undefined && priceCell.v !== null ? parseFloat(priceCell.v) : null;
    if (id && v !== null && !isNaN(v)) {
        productPrices[id] = v;
    }
}

const uniquePrices = [...new Set(Object.values(productPrices))].sort((a, b) => b - a);
console.log('商品数量:', Object.keys(productPrices).length);
console.log('去重价格数:', uniquePrices.length);
console.log('所有去重价格（降序）:', uniquePrices);

const topThirdCount = Math.ceil(uniquePrices.length / 3);
const threshold = uniquePrices[topThirdCount - 1];
console.log('\n前1/3数量:', topThirdCount);
console.log('阈值（第' + topThirdCount + '大的值）:', threshold);
console.log('阈值在去重数组中的位置:', uniquePrices.indexOf(threshold));

// 哪些商品在top 1/3
const topProducts = Object.entries(productPrices).filter(([id, p]) => p >= threshold);
console.log('\n价格在top 1/3的商品数:', topProducts.length);
for (const [id, p] of topProducts) {
    console.log('  ' + id + ': ' + p);
}
