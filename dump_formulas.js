import ExcelJS from 'exceljs';

async function dump() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('./Hasil Otomasi.xlsx');
    const sheet = workbook.getWorksheet(1);

    // Rows 16 to 30
    // Cols D(4) E(5) F(6) G(7) H(8)

    for (let r = 16; r <= 30; r++) {
        for (let c = 4; c <= 8; c++) {
            const cell = sheet.getCell(r, c);
            if (cell.value) {
                const head = `R${r}C${c} [${["", "A", "B", "C", "D", "E", "F", "G", "H"][c]}]:`;
                const val = cell.formula ? `FORMULA=${cell.formula}` : `VALUE=${cell.value}`;
                console.log(`${head} ${val}`);
            }
        }
    }
}
dump();
