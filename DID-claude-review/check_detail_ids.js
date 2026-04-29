const XLSX = require('xlsx');
const fs = require('fs');

const qaPath = './问答.xlsx';
const commentPath = './评论数.xlsx';
const detailPath = './商品详情.xlsx';

function getIdsFromColumn(workbook, colIndex) {
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const ids = new Set();
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: row, c: colIndex })];
        if (cell && cell.v !== null && cell.v !== undefined) {
            ids.add(cell.v);
        }
    }
    return ids;
}

const wb1 = XLSX.readFile(qaPath, { cellNF: true, cellDates: true });
const wb2 = XLSX.readFile(commentPath, { cellNF: true, cellDates: true });
const wb3 = XLSX.readFile(detailPath, { cellNF: true, cellDates: true });

const qaIds = getIdsFromColumn(wb1, 0);
const commentIds = getIdsFromColumn(wb2, 0);
const detailIds = getIdsFromColumn(wb3, 0);

const onlyInQaComment = new Set([...qaIds].filter(x => !detailIds.has(x)));
const onlyInDetail = new Set([...detailIds].filter(x => !qaIds.has(x)));

let result = '';
result += `商品详情.xlsx ID count: ${detailIds.size}\n`;
result += `问答.xlsx ID count: ${qaIds.size}\n`;
result += `评论数.xlsx ID count: ${commentIds.size}\n\n`;

result += `详情表包含问答表所有ID? ${detailIds.size >= qaIds.size && onlyInQaComment.size === 0 ? 'YES' : 'NO'}\n`;
result += `详情表包含评论表所有ID? ${detailIds.size >= commentIds.size && onlyInDetail.size === 0 ? 'YES' : 'NO'}\n\n`;

result += `详情表中有但问答表中没有的ID数: ${onlyInDetail.size}\n`;
result += `问答表中有但详情表中没有的ID数: ${onlyInQaComment.size}\n`;

if (onlyInQaComment.size > 0) {
    result += `问答表中有但详情表中没有的ID: ${JSON.stringify([...onlyInQaComment])}\n`;
}
if (onlyInDetail.size > 0) {
    result += `详情表中有但问答表中没有的ID: ${JSON.stringify([...onlyInDetail].slice(0, 30))}${onlyInDetail.size > 30 ? ' ...' : ''}\n`;
}

fs.writeFileSync('./detail_id_check_result.txt', result, 'utf-8');
console.log(result);
