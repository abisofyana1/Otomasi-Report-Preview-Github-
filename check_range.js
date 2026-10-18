import ExcelJS from 'exceljs';

async function checkRange() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('./Hasil Otomasi.xlsx');
    const sheet = workbook.getWorksheet(1);

    console.log("Checking D16:H27...");
    for (let r = 16; r <= 27; r++) {
        let line = `Row ${r}: `;
        for (let c = 4; c <= 8; c++) { // D=4, H=8
            const cell = sheet.getCell(r, c);
            const val = cell.value;
            // Value might be an object if formula
            let show = "";
            if (val && typeof val === 'object' && val.formula) show = `(Formula: ${val.formula})`;
            else show = `"${val}"`;

            line += `| Col${c}: ${show} `;
        }
        console.log(line);
    }
}
checkRange();
