const XLSX = require('xlsx');

const wb = XLSX.readFile('./问答_合并商品信息_清洗后.xlsx', { cellNF: true, cellDates: true });
const sheet = wb.Sheets[wb.SheetNames[0]];

// 写入CSV
XLSX.writeFile(wb, './问答_合并商品信息_清洗后.csv');
console.log('已保存为CSV: 问答_合并商品信息_清洗后.csv');
