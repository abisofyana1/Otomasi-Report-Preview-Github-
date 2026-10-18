import XLSX from 'xlsx';
import fs from 'fs';

const SOURCE_FILE = './Test 1.xlsx';
const TEMPLATE_FILE = './Hasil Otomasi.xlsx';
const OUTPUT_FILE = './Test 1_Result.xlsx';

const TARGET_AHT_MAP = {
    'SAT': 300,
    'T&R': 300,
    'T & R': 300, // Handle variations
    'BS': 300,
    'EMERALD': 390,
    'BK': 240,
    'KK': 240
};

// 1. Read Template for Summary Block
const wbTemplate = XLSX.readFile(TEMPLATE_FILE);
const sheetTemplate = wbTemplate.Sheets[wbTemplate.SheetNames[0]];

const summaryBlock = [];
const tStartR = 15; // Row 16
const tEndR = 26;   // Row 27
const tStartC = 3;  // Col D
const tEndC = 7;    // Col H

// Original Data End Row (deduced from context or formulas) => 14
const origDataEndRow = 14;

for (let r = tStartR; r <= tEndR; r++) {
    for (let c = tStartC; c <= tEndC; c++) {
        const cell = sheetTemplate[XLSX.utils.encode_cell({ r, c })];
        if (cell) {
            summaryBlock.push({
                rOffset: r - tStartR, // 0 to 11
                cOffset: c - tStartC, // 0 to 4
                cell: cell
            });
        }
    }
}

// 2. Read Source File
const wbSource = XLSX.readFile(SOURCE_FILE);
const sheetSource = wbSource.Sheets[wbSource.SheetNames[0]];
const sourceJson = XLSX.utils.sheet_to_json(sheetSource, { header: 1 });
const headers = sourceJson[0];
const idxSkill = headers.findIndex(h => h && h.toString().toLowerCase().trim() === 'skill');

if (idxSkill === -1) {
    console.error("No 'Skill' column found!");
    process.exit(1);
}

// Group Data
const groups = {};
for (let i = 1; i < sourceJson.length; i++) {
    const row = sourceJson[i];
    if (!row || row.length === 0) continue;

    // Normalize logic for Target AHT
    const skillVal = row[idxSkill];
    const skillKey = skillVal ? skillVal.toString().trim() : 'Uncategorized';

    // Find matching target
    // Check direct match or partial? Assuming direct map keys match valid skills
    let targetAHT = 0;
    // naive match
    const upperSkill = skillKey.toUpperCase();
    if (TARGET_AHT_MAP[upperSkill]) targetAHT = TARGET_AHT_MAP[upperSkill];
    // try removing spaces for 'T & R' -> 'T&R' check
    else if (TARGET_AHT_MAP[upperSkill.replace(/\s/g, '')]) targetAHT = TARGET_AHT_MAP[upperSkill.replace(/\s/g, '')];

    // Inject Target AHT into Column F (Index 5)
    // Ensure row has enough columns
    while (row.length <= 5) row.push(null);
    row[5] = targetAHT;

    if (!groups[skillKey]) groups[skillKey] = [];
    groups[skillKey].push(row);
}

const newWb = XLSX.utils.book_new();

// 3. Process each group
Object.keys(groups).forEach(skill => {
    const rows = groups[skill];
    // Create sheet with headers + data
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

    // Total rows (1 header + N data)
    const dataCount = rows.length;
    const lastDataRowHeaderIncluded = dataCount + 1; // 1-based, last row index
    // e.g. if 7 rows, header=1, data=2..8. LastDataRow = 8.

    // Apply Data Row Formulas (Col A and G)
    for (let R = 1; R <= dataCount; ++R) { // 0-based index relative to data start (row 2)
        const trueRow = R + 1; // 1-based Excel Row
        // Col A (No) -> ROWS(A$2:A{trueRow})
        ws[XLSX.utils.encode_cell({ r: trueRow - 1, c: 0 })] = { t: 'n', f: `ROWS(A$2:A${trueRow})` };

        // Col G (Sesuai AHT) -> IF(E{r}<=F{r}...)
        // Preserve the Target AHT value in F we set earlier? Yes, it's in `rows` array.
        // But AOToSheet puts values, not formulas. We overwrite for col G.
        ws[XLSX.utils.encode_cell({ r: trueRow - 1, c: 6 })] = { t: 's', f: `IF(E${trueRow}<=F${trueRow},"Target","Tidak Target")` };
    }

    // 4. Append Summary Block
    // Starting row: Leave 1 empty row after last data? Original was: Data end 14, Summ start 16. Gap 1.
    const gap = 1;
    const summaryStartRowInfo = dataCount + 1 + gap; // 1-based index of where summary starts? No, logic:
    // If data ends at row 8 (index 7).
    // Gap row is 9 (index 8).
    // Summary starts row 10 (index 9).
    const range = XLSX.utils.decode_range(ws['!ref']);
    const newSummaryStartR = range.e.r + 2; // e.r is last data index. +1 is empty row, +2 is summary start.

    // Calculate shift offset for formulas
    const rowOffset = newSummaryStartR - tStartR; // Delta between new pos and old pos (15)

    summaryBlock.forEach(item => {
        const targetR = newSummaryStartR + item.rOffset;
        const targetC = tStartC + item.cOffset; // Same columns D-H

        // Clone cell object
        const newCell = { ...item.cell };

        // Adjust Formula
        if (newCell.f) {
            // Regex to find row numbers
            newCell.f = newCell.f.replace(/([A-Z]+)([0-9]+)/g, (match, col, rowStr) => {
                const r = parseInt(rowStr);
                // Case 1: Reference to Data Range (<= origDataEndRow)
                if (r <= origDataEndRow) {
                    // Start row usually 2, stays 2.
                    // End row 14 becomes lastDataRowHeaderIncluded.
                    if (r === origDataEndRow) return `${col}${lastDataRowHeaderIncluded}`;
                    if (r === 2) return `${col}2`;
                    // If it refers to middle of data, tricky. Usually ranges are X2:X14.
                    return match;
                }
                // Case 2: Reference to Summary Range (>= tStartR + 1 = 16)
                if (r >= tStartR + 1) {
                    const newR = r + rowOffset;
                    return `${col}${newR}`;
                }
                return match;
            });
        }

        ws[XLSX.utils.encode_cell({ r: targetR, c: targetC })] = newCell;
    });

    // Update Sheet Range
    const totalRows = newSummaryStartR + (tEndR - tStartR) + 1;
    const newRange = { s: { r: 0, c: 0 }, e: { r: totalRows, c: headers.length > tEndC ? headers.length : tEndC + 1 } }; // Safe bound
    ws['!ref'] = XLSX.utils.encode_range(newRange);

    // Append Sheet
    let safeName = skill.replace(/[\*\?\/\[\]\\\: ]/g, ' ').trim().substring(0, 31);
    let uniqueName = safeName;
    let counter = 1;
    while (newWb.SheetNames.includes(uniqueName)) {
        uniqueName = `${safeName.substring(0, 28)}(${counter})`;
        counter++;
    }

    XLSX.utils.book_append_sheet(newWb, ws, uniqueName);
});

XLSX.writeFile(newWb, OUTPUT_FILE);
console.log("Done.");
