import ExcelJS from 'exceljs';

async function checkLabels() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('./Hasil Otomasi.xlsx');
    const sheet = workbook.getWorksheet(1);

    console.log("--- Checking Labels in D16:H30 ---");
    for (let r = 16; r <= 30; r++) {
        const cell = sheet.getCell(`D${r}`); // Check Col D for labels
        console.log(`Row ${r}: ${cell.value}`);
    }
}
checkLabels();
