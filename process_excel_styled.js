import ExcelJS from 'exceljs';
import fs from 'fs';

const SOURCE_FILE = './Test 1.xlsx';
const TEMPLATE_FILE = './Hasil Otomasi.xlsx';
const OUTPUT_FILE = './Test 1_Result.xlsx';

const TARGET_AHT_MAP = {
    'SAT': 300, 'T&R': 300, 'T & R': 300, 'BS': 300,
    'EMERALD': 390, 'BK': 240, 'KK': 240
};

async function main() {
    // 1. Read Template for Summary Block (using the range D16:H27 which covers headers up to 'Menit')
    const workbookTemplate = new ExcelJS.Workbook();
    await workbookTemplate.xlsx.readFile(TEMPLATE_FILE);
    const sheetTemplate = workbookTemplate.getWorksheet(1);

    // We will capture styles and values/formulas from D16:H27
    // Based on check_labels: 
    // R19: Total Agent, R21: % Target Agent, R23: % Ready... R27: Menit.
    // So D16:H27 seems to cover the relevant rows!

    const summaryData = [];
    const tStartR = 16;
    const tEndR = 27;
    const tStartC = 4; // D
    const tEndC = 8;   // H

    for (let r = tStartR; r <= tEndR; r++) {
        const rowData = [];
        for (let c = tStartC; c <= tEndC; c++) {
            const cell = sheetTemplate.getCell(r, c);
            rowData.push({
                value: cell.value,
                style: cell.style, // Capture style
                formula: cell.formula,
                result: cell.result,
                model: cell.model // capture full model for complex stuff
            });
        }
        summaryData.push(rowData);
    }

    // 2. Read Source
    const workbookSource = new ExcelJS.Workbook();
    await workbookSource.xlsx.readFile(SOURCE_FILE);
    const sheetSource = workbookSource.getWorksheet(1);

    // Extract Data
    // Assuming standard table: Header row 1, data starts row 2
    const sourceRows = [];
    // Get headers
    const headerRow = sheetSource.getRow(1);
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        headers[colNumber] = cell.value;
    });

    // Find skill col
    let skillColIdx = -1;
    for (let i = 1; i < headers.length; i++) {
        if (headers[i] && headers[i].toString().toLowerCase().trim() === 'skill') {
            skillColIdx = i;
            break;
        }
    }

    if (skillColIdx === -1) {
        console.error("Skill column not found");
        return;
    }

    sheetSource.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // skip header
        const rowVals = [];
        // Extract cell values carefully
        // We know structure approx 8-9 cols
        for (let c = 1; c <= Math.max(headers.length, 9); c++) {
            rowVals[c] = row.getCell(c).value;
        }
        sourceRows.push(rowVals);
    });

    // Group
    const groups = {};
    sourceRows.forEach(row => {
        const skill = row[skillColIdx] ? row[skillColIdx].toString().trim() : "Uncategorized";
        if (!groups[skill]) groups[skill] = [];
        groups[skill].push(row);
    });

    // 3. Create Output
    const wbOut = new ExcelJS.Workbook();

    for (const skill of Object.keys(groups)) {
        const rows = groups[skill];

        // Output Sheet Name
        let sheetName = skill.replace(/[\*\?\/\[\]\\\: ]/g, ' ').trim().substring(0, 31);
        // Ensure unique
        let counter = 1;
        let baseName = sheetName;
        while (wbOut.getWorksheet(sheetName)) {
            sheetName = `${baseName.substring(0, 28)}(${counter})`;
            counter++;
        }

        const ws = wbOut.addWorksheet(sheetName);

        // Define Columns / Headers
        // Headers are: No, ID Oceana, Agent Name, Skill, AHT by ACD, Target AHT, Sesuai AHT, % Ready
        // Indices: 1, 2, 3, 4, 5, 6, 7, 8

        // Set Headers
        // We can manually set headers to ensure exact names
        const headerValues = [
            "No", "ID Oceana", "Agent Name", "Skill", "AHT by ACD",
            "Target AHT", "Sesuai AHT", "Persentase" // Changed % Ready to Persentase per request? 
            // "Kolom %Ready ubah menjadi Persentase" -> Label change.
        ];

        // The user said: "Kolom dan isi ... %Ready tolong posisikan dicenter".
        // And "Kolom %Ready ubah menjadi Persentase".
        // I will change the Header Text to "Persentase".

        const rowHeader = ws.getRow(1);
        for (let i = 0; i < headerValues.length; i++) {
            const cell = rowHeader.getCell(i + 1);
            cell.value = headerValues[i];
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            cell.font = { bold: true };
        }

        // Add Data
        rows.forEach((rowVals, index) => {
            const r = index + 2;
            const dataRow = ws.getRow(r);

            // Col A: No -> ROWS(A$2:Ar)
            dataRow.getCell(1).value = { formula: `ROWS(A$2:A${r})` };

            // Col B: ID -> rowVals[2]
            dataRow.getCell(2).value = rowVals[2];

            // Col C: Name -> rowVals[3]
            dataRow.getCell(3).value = rowVals[3];

            // Col D: Skill -> rowVals[4]
            dataRow.getCell(4).value = rowVals[4];

            // Col E: AHT by ACD -> rowVals[5]
            dataRow.getCell(5).value = rowVals[5];

            // Col F: Target AHT (Map)
            let targetAHT = 0;
            const sKey = rowVals[4] ? rowVals[4].toString().toUpperCase() : "";
            if (TARGET_AHT_MAP[sKey]) targetAHT = TARGET_AHT_MAP[sKey];
            else if (TARGET_AHT_MAP[sKey.replace(/\s/g, '')]) targetAHT = TARGET_AHT_MAP[sKey.replace(/\s/g, '')];

            dataRow.getCell(6).value = targetAHT;

            // Col G: Sesuai AHT -> IF(E<=F...)
            // Comparing AHT(E) and Target(F)
            dataRow.getCell(7).value = { formula: `IF(E${r}<=F${r},"Target","Tidak Target")` };

            // Col H: % Ready -> rowVals[8] (was 8 in source? check source headers again or just assume idx 8)
            // Wait, source array indices match column numbers if I used rowVals[c] = cell.value
            // So if source has "Skill" at 4, AHT at 5. Target was NOT in source?
            // "Test 1.xlsx" structure:
            // No, ID, Name, Skill, AHT, Target?, Sesuai?, %Ready?
            // Let's rely on data being copied.
            // If source didn't have Target/Sesuai, then %Ready might be at a different index.
            // Let's inspect headers of source again.
            // check_headers.js said: ["No","ID Oceana","Agent Name","Skill","AHT by ACD","Target\r\nAHT","Sesuai\r\nAHT","% Ready"]
            // So indices: 1,2,3,4,5,6,7,8.
            // So %Ready is 8.
            dataRow.getCell(8).value = rowVals[8];

            // Apply Styling
            // 1. All Border
            for (let c = 1; c <= 8; c++) {
                const cell = dataRow.getCell(c);
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

                // 2. Alignment Centering
                // Valid Cols for center: 1(No), 2(ID), 4(Skill), 5(AHT), 6(Target), 7(Sesuai), 8(%Ready)
                if ([1, 2, 4, 5, 6, 7, 8].includes(c)) {
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                }

                // 3. Format %Ready (Col 8)
                if (c === 8) {
                    cell.numFmt = '0.00%'; // "Misalnya : 100,00%"
                }
            }
        });

        // Summary Block
        const lastDataRow = rows.length + 1;
        const summaryStartRow = lastDataRow + 2; // Gap 1

        const rowOffset = summaryStartRow - tStartR; // Delta

        for (let r = 0; r < summaryData.length; r++) {
            const targetR = summaryStartRow + r;
            const srcRowData = summaryData[r]; // array of cells
            const templateR = tStartR + r;

            for (let c = 0; c < srcRowData.length; c++) {
                const srcCell = srcRowData[c];
                const targetC = tStartC + c; // D is 4.

                const targetCell = ws.getCell(targetR, targetC);

                // Copy Value/Formula
                if (srcCell.model && srcCell.model.formula) {
                    // Adjust formula
                    let f = srcCell.model.formula;
                    // Replace Ranges
                    f = f.replace(/([A-Z]+)([0-9]+)/g, (match, colStr, rowStr) => {
                        const refR = parseInt(rowStr);
                        // If ref refers to Data (Row 2 .. 14 in template)
                        if (refR >= 2 && refR <= 14) {
                            if (refR === 2) return `${colStr}2`;
                            if (refR === 14) return `${colStr}${lastDataRow}`;
                            return match; // Shouldn't happen purely for total ranges
                        }
                        // If ref refers to Summary
                        if (refR >= 16) {
                            return `${colStr}${refR + rowOffset}`;
                        }
                        return match;
                    });
                    targetCell.value = { formula: f };
                } else {
                    targetCell.value = srcCell.value;
                }

                // Styles: Copy Border
                // Or just apply All Borders request? "Untuk semua kolom dan isi tolong All Border."
                // Usually implies the data table + summary table.
                targetCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

                // Specific Formats based on Label Logic
                // We check the Label in Column D (Index 0 of srcRowData)
                const labelVal = summaryData[r][0].value;

                // If this is the Label cell (Col D)
                if (targetC === 4) {
                    // Do nothing special unless label itself needs formatting
                }

                // If this is the Value cell (Col E? Index 1 of srcRowData?)
                // check_labels said R21 "% Target Agent".
                // If Label is "% Target Agent" (R21), then Value is likely E21.
                // If Label is "% Ready" (R23), Value E23.
                // If Label is "Rata Rata Handeling" (R25), Value E25.
                // If "Menit / Detik" or "Menit" (R26/27).

                // We are iterating columns D..H.
                // If targetC is 5 (Col E) and labelVal looks like...
                if (targetC === 5 && labelVal) {
                    const labelStr = labelVal.toString();

                    if (labelStr.includes("% Target Agent") || labelStr.includes("% Ready")) {
                        targetCell.numFmt = '0.00%';
                    }
                    if (labelStr.includes("Rata Rata Handeling")) {
                        targetCell.numFmt = '0'; // No decimals
                    }
                    if (labelStr.includes("Menit")) {
                        targetCell.numFmt = 'mm:ss'; // "MM:SS"
                    }
                }
            }
        }

    }

    await wbOut.xlsx.writeFile(OUTPUT_FILE);
    console.log("Processing Complete.");
}

main().catch(console.error);
