import ExcelJS from 'exceljs';

async function checkFormulas() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('./Hasil Otomasi.xlsx');
    const sheet = workbook.getWorksheet(1);

    // We are interested in D16:H27
    // D is col 4. H is col 8.

    const rows = [];
    for (let r = 16; r <= 27; r++) {
        const row = { row: r };
        for (let c = 4; c <= 8; c++) {
            const cell = sheet.getCell(r, c);
            if (cell.value) {
                // key by column letter D, E, F, G, H
                const colLet = ["", "A", "B", "C", "D", "E", "F", "G", "H"][c];
                row[colLet] = cell.formula ? `=${cell.formula}` : cell.value;
            }
        }
        rows.push(row);
    }
    console.table(rows);
    console.log(JSON.stringify(rows, null, 2));
}

checkFormulas();
