import XLSX from 'xlsx';

const wb = XLSX.readFile('./Hasil Otomasi.xlsx');
const sheet = wb.Sheets[wb.SheetNames[0]];

const startR = 15; // Row 16
const endR = 26;   // Row 27
const startC = 3;  // Col D
const endC = 7;    // Col H

for (let r = startR; r <= endR; r++) {
    let rowStr = `R${r + 1}: `;
    for (let c = startC; c <= endC; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })];
        const val = cell ? (cell.f ? `[F=${cell.f}]` : cell.v) : "";
        rowStr += ` | ${val}`;
    }
    console.log(rowStr);
}
