const XLSX = require('xlsx');
const fs = require('fs');

const qaPath = './问答.xlsx';
const commentPath = './评论数.xlsx';
const detailPath = './商品详情.xlsx';

function getSheetInfo(workbook) {
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const totalRows = range.e.r - range.s.r + 1;
    const dataRows = range.e.r - range.s.r;
    let nonEmptyRows = 0;
    for (let row = range.s.r; row <= range.e.r; row++) {
        let hasData = false;
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
            if (cell && cell.v !== null && cell.v !== undefined && cell.v !== '') {
                hasData = true;
                break;
            }
        }
        if (hasData) nonEmptyRows++;
    }
    return { totalRows, dataRows, nonEmptyRows };
}

const wb1 = XLSX.readFile(qaPath, { cellNF: true, cellDates: true });
const wb2 = XLSX.readFile(commentPath, { cellNF: true, cellDates: true });
const wb3 = XLSX.readFile(detailPath, { cellNF: true, cellDates: true });

const info1 = getSheetInfo(wb1);
const info2 = getSheetInfo(wb2);
const info3 = getSheetInfo(wb3);

let result = '';
result += `问答.xlsx: total_rows=${info1.totalRows}, data_rows=${info1.dataRows}, non_empty=${info1.nonEmptyRows}\n`;
result += `评论数.xlsx: total_rows=${info2.totalRows}, data_rows=${info2.dataRows}, non_empty=${info2.nonEmptyRows}\n`;
result += `商品详情.xlsx: total_rows=${info3.totalRows}, data_rows=${info3.dataRows}, non_empty=${info3.nonEmptyRows}\n`;

fs.writeFileSync('./row_count_result.txt', result, 'utf-8');
console.log(result);
