import ExcelJS from 'exceljs';

async function scanTemplate() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('./Hasil Otomasi.xlsx');
    const sheet = workbook.getWorksheet(1); // 1st sheet

    console.log("Scanning Summary Block (Rows 16-30 approx)...");

    // Scan locally around the known area
    for (let r = 15; r <= 30; r++) {
        const row = sheet.getRow(r);
        let rowStr = `R${r}: `;
        row.eachCell((cell, colNumber) => {
            rowStr += `[C${colNumber}: ${cell.value}] `;
        });
        console.log(rowStr);
    }
}

scanTemplate();
