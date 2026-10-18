import XLSX from 'xlsx';

const wb = XLSX.readFile('./Hasil Otomasi.xlsx');
const sheet = wb.Sheets[wb.SheetNames[0]];

// Target range D16:H27 (0-indexed: R15:C3 to R26:C7)
// D=3, E=4, F=5, G=6, H=7
const startR = 15; // Row 16
const endR = 26;   // Row 27
const startC = 3;  // Col D
const endC = 7;    // Col H

console.log("--- Extracting Summary Section (D16:H27) ---");
const summary = [];

for (let r = startR; r <= endR; r++) {
    const rowData = [];
    for (let c = startC; c <= endC; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })];
        rowData.push({
            r, c,
            v: cell ? cell.v : null,
            f: cell ? cell.f : null,
            t: cell ? cell.t : null
        });
    }
    summary.push(rowData);
}

console.log(JSON.stringify(summary, null, 2));
