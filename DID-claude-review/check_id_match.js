const XLSX = require('xlsx');

const qaPath = './问答.xlsx';
const commentPath = './评论数.xlsx';
const outPath = './id_check_result.txt';
const fs = require('fs');

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

const qaIds = getIdsFromColumn(wb1, 0);
const commentIds = getIdsFromColumn(wb2, 0);

const onlyInQa = new Set([...qaIds].filter(x => !commentIds.has(x)));
const onlyInComment = new Set([...commentIds].filter(x => !qaIds.has(x)));
const intersection = new Set([...qaIds].filter(x => commentIds.has(x)));

let result = '';
result += `问答.xlsx ID count: ${qaIds.size}\n`;
result += `评论数.xlsx ID count: ${commentIds.size}\n`;
result += `Intersection: ${intersection.size}\n`;
result += `Only in 问答.xlsx: ${onlyInQa.size}\n`;
result += `Only in 评论数.xlsx: ${onlyInComment.size}\n`;

if (onlyInQa.size > 0) {
    const arr = [...onlyInQa].slice(0, 30);
    result += `Only in 问答.xlsx IDs: ${JSON.stringify(arr)}\n`;
}
if (onlyInComment.size > 0) {
    const arr = [...onlyInComment].slice(0, 30);
    result += `Only in 评论数.xlsx IDs: ${JSON.stringify(arr)}\n`;
}

if (onlyInQa.size === 0 && onlyInComment.size === 0) {
    result += 'Result: IDs match perfectly\n';
} else {
    result += 'Result: IDs DO NOT match perfectly\n';
}

fs.writeFileSync(outPath, result, 'utf-8');
console.log(result);
