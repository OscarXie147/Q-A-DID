const XLSX = require('xlsx');

const wb = XLSX.readFile('./商品详情_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];
const range = XLSX.utils.decode_range(sheet['!ref']);

const headers = [];
for (let col = range.s.c; col <= range.e.c; col++) {
    const c = sheet[XLSX.utils.encode_cell({ r: 0, c: col })];
    headers.push(c && c.v != null ? String(c.v) : 'col_' + col);
}
const baiduIdx = headers.indexOf('百度指数');
const awarenessIdx = headers.indexOf('brand_awareness');
const brandIdx = headers.indexOf('品牌');

const targetBrands = ['悦美时刻', 'ALLSMILE/哎哟咪', 'Girlcult', 'KATO', '娜斯', '媞妃特', '彭世', 'CECILE MAIA', '简初'];

console.log('检查9个品牌的当前值:');
for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const br = sheet[XLSX.utils.encode_cell({ r: row, c: brandIdx })];
    const brand = br && br.v != null ? String(br.v).trim() : '';
    if (targetBrands.some(b => brand.includes(b))) {
        const bi = sheet[XLSX.utils.encode_cell({ r: row, c: baiduIdx })];
        const ai = sheet[XLSX.utils.encode_cell({ r: row, c: awarenessIdx })];
        console.log('Row' + (row+1) + ' brand=' + brand + ' 百度指数=' + JSON.stringify(bi ? bi.v : 'undefined') + ' awareness=' + JSON.stringify(ai ? ai.v : 'undefined'));
    }
}
