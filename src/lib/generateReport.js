import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const TARGET_AHT_MAP = {
    'SAT': 300,
    'T&R': 300, 'T & R': 300,
    'BS': 300,
    'EMERALD': 390,
    'BK': 240,
    'KK': 240,
    'BPC': 270,
    'MC': 270
};

const ALLOWED_SKILLS = ['BK', 'KK', 'BPC', 'EMERALD', 'MC', 'SAT', 'T&R', 'BS'];

// Helper to normalize headers
const norm = (h) => h ? h.toString().toLowerCase().replace(/[^a-z0-9]/g, '') : '';

// ---------------------------------------------------------------------------
// Helper: Get Numeric Safe Value (Strict Number Enforcement)
// Prevents XML Errors by ensuring metrics are written as Numbers, not Strings
// ---------------------------------------------------------------------------
const getNumericSafeVal = (row, colIdx) => {
    try {
        if (!colIdx || colIdx < 1) return 0;
        const cell = row.getCell(colIdx);
        let val = cell.value;

        // 1. Handle Objects
        if (typeof val === 'object') {
            if (val === null) return 0;
            if (val instanceof Date) return val; // Allow Date/Time objects

            // Unwrap
            if (val.result !== undefined) val = val.result;
            else if (val.text !== undefined) val = val.text;
            else if (val.richText) val = val.richText.map(t => t.text).join('');
        }

        // 2. Parse Number
        if (typeof val === 'number') {
            return Number.isFinite(val) ? val : 0;
        }

        if (typeof val === 'string') {
            // Remove commas, percent signs, trim
            const clean = val.replace(/,/g, '').replace(/%/g, '').trim();
            if (clean === '') return 0;
            const parsed = parseFloat(clean);

            // Handle Percentages (if string was '100', it's 100. If '100%', it's 100 or 1? Logic depends on Excel)
            // Usually if reading raw value: 0.95 -> 95%.
            // If reading string "95%", parseFloat is 95. 
            // We'll trust parseFloat.
            return Number.isFinite(parsed) ? parsed : 0;
        }

        return 0;
    } catch {
        return 0;
    }
};

// ---------------------------------------------------------------------------
// Helper: Get Safe Value (Crucial for preventing Excel Corruption & Error 500)
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Helper: Get Safe Value (Paranoid Mode)
// ---------------------------------------------------------------------------
const getSafeVal = (row, colIdx) => {
    try {
        if (!colIdx || colIdx < 1) return 0;
        const cell = row.getCell(colIdx);
        let val = cell.value;

        // 1. Handle Null/Undefined
        if (val === null || val === undefined) return 0;

        // 2. Handle Objects (Formulas, RichText, Hyperlinks)
        if (typeof val === 'object') {
            if (val instanceof Date) return val; // Dates are OK

            // Unpack Formula Result
            if (val.hasOwnProperty('result')) {
                // If result is Error object, return 0
                if (val.result && typeof val.result === 'object') return 0;
                val = val.result;
            }
            // Unpack Rich Text
            else if (val.hasOwnProperty('richText') && Array.isArray(val.richText)) {
                val = val.richText.map(rt => rt.text).join('');
            }
            // Unpack Hyperlink / Text Object
            else if (val.hasOwnProperty('text')) {
                val = val.text;
            }
            else {
                // Unknown Object? Block it.
                return 0;
            }
        }

        // 3. Re-Check Value after Unpacking
        if (val === null || val === undefined) return 0;

        // 4. Handle Numbers (NaN, Infinity)
        if (typeof val === 'number') {
            if (!Number.isFinite(val)) return 0;
            return val;
        }

        // 5. Handle Strings (Trim)
        if (typeof val === 'string') {
            return val.trim();
        }

        // 6. Fallback
        return val;

    } catch (err) {
        // Absolute Safety Net
        return 0;
    }
};

// ---------------------------------------------------------------------------
// Helper: Parse Day from Filename
// ---------------------------------------------------------------------------
const getDayFromFileName = (name) => {
    // strict "1.xlsx"
    const exactMatch = name.match(/^(\d+)\.xlsx$/i);
    if (exactMatch) {
        const num = parseInt(exactMatch[1], 10);
        if (num >= 1 && num <= 31) return num;
    }
    // embedded "Report 31 Jan.xlsx"
    const match = name.match(/(\d+)/);
    if (match) {
        const num = parseInt(match[1], 10);
        if (num >= 1 && num <= 31) return num;
    }
    return null;
};

// ---------------------------------------------------------------------------
// 1. Generate Preview Data
// ---------------------------------------------------------------------------
export const generatePreviewData = (originalData) => {
    if (!originalData || originalData.length === 0) return [];

    const inputHeaders = originalData[0];
    const rows = originalData.slice(1);

    const map = {
        id: inputHeaders.findIndex(h => norm(h).includes('idoceana')),
        name: inputHeaders.findIndex(h => norm(h).includes('agentname')),
        skill: inputHeaders.findIndex(h => norm(h) === 'skill'),
        aht: inputHeaders.findIndex(h => norm(h).includes('ahtbyacd')),
        pct: inputHeaders.findIndex(h => norm(h).includes('ready') || norm(h).includes('%ready'))
    };

    if (map.skill === -1) map.skill = 3;
    if (map.id === -1) map.id = 1;
    if (map.name === -1) map.name = 2;
    if (map.aht === -1) map.aht = 4;

    const previewHeaders = ["No", "ID Oceana", "Agent Name", "Skill", "AHT by ACD", "Target AHT", "Sesuai AHT", "%Ready"];
    const previewRows = [];
    let noCounter = 1;

    rows.forEach((row) => {
        const skill = row[map.skill] ? row[map.skill].toString().trim() : 'Uncategorized';
        const normSkill = skill.toUpperCase().replace(/\s/g, '');
        const isAllowed = ALLOWED_SKILLS.some(allowed => allowed.replace(/\s/g, '') === normSkill);

        if (!isAllowed) return;

        const idVal = row[map.id];
        const nameVal = row[map.name];
        const ahtVal = Number(row[map.aht]) || 0;

        let targetAHT = 0;
        const skillKey = skill.toUpperCase();
        if (TARGET_AHT_MAP[skillKey]) targetAHT = TARGET_AHT_MAP[skillKey];
        else if (TARGET_AHT_MAP[skillKey.replace(/\s/g, '')]) targetAHT = TARGET_AHT_MAP[skillKey.replace(/\s/g, '')];

        const sesuai = ahtVal <= targetAHT ? "Target" : "Tidak Target";

        let valPct = map.pct !== -1 ? row[map.pct] : null;
        if (typeof valPct === 'number' && valPct > 1) valPct = valPct / 100;
        else if (typeof valPct === 'string') {
            let parsed = parseFloat(valPct.replace('%', ''));
            if (!isNaN(parsed)) {
                if (parsed > 1) parsed = parsed / 100;
                valPct = parsed;
            }
        }
        const pctDisplay = valPct != null ? (typeof valPct === 'number' ? (valPct * 100).toFixed(2) + '%' : valPct) : '';

        previewRows.push([noCounter++, idVal, nameVal, skill, ahtVal, targetAHT, sesuai, pctDisplay]);
    });

    return [previewHeaders, ...previewRows];
};

// ---------------------------------------------------------------------------
// 2. Generate AHT Report (Excel)
// ---------------------------------------------------------------------------
export const generateAHTReport = async (originalData, originalFileName) => {
    if (!originalData || originalData.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const inputHeaders = originalData[0];
    const rows = originalData.slice(1);

    // ... (Keeping existing implementation for AHT to save space, assuming it works fine) ...
    // Note: Re-implementing AHT part since we are overwriting file.

    const map = {
        id: inputHeaders.findIndex(h => norm(h).includes('idoceana')),
        name: inputHeaders.findIndex(h => norm(h).includes('agentname')),
        skill: inputHeaders.findIndex(h => norm(h) === 'skill'),
        aht: inputHeaders.findIndex(h => norm(h).includes('ahtbyacd')),
        pct: inputHeaders.findIndex(h => norm(h).includes('ready') || norm(h).includes('%ready'))
    };

    if (map.skill === -1) map.skill = 3;
    if (map.id === -1) map.id = 1;
    if (map.name === -1) map.name = 2;
    if (map.aht === -1) map.aht = 4;

    const groups = {};
    rows.forEach(row => {
        const skill = row[map.skill] ? row[map.skill].toString().trim() : 'Uncategorized';
        if (!groups[skill]) groups[skill] = [];
        groups[skill].push(row);
    });

    for (const skill of Object.keys(groups)) {
        const normSkill = skill.toUpperCase().replace(/\s/g, '');
        const isAllowed = ALLOWED_SKILLS.some(allowed => allowed.replace(/\s/g, '') === normSkill);
        if (!isAllowed) continue;

        // ... (standard AHT generation logic) ...
        const groupRows = groups[skill];
        let sheetName = skill.replace(/[\*\?\/\[\]\\\: ]/g, ' ').trim().substring(0, 31);
        let counter = 1;
        let baseName = sheetName;
        while (workbook.getWorksheet(sheetName)) {
            sheetName = `${baseName.substring(0, 28)}(${counter})`;
            counter++;
        }
        const ws = workbook.addWorksheet(sheetName);
        ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

        const reportHeaders = ["No", "ID Oceana", "Agent Name", "Skill", "AHT by ACD", "Target AHT", "Sesuai AHT", "%Ready"];
        const headerRow = ws.getRow(1);
        reportHeaders.forEach((h, i) => {
            const cell = headerRow.getCell(i + 1);
            cell.value = h;
            cell.font = { bold: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' } };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });

        let totalTarget = 0;
        let totalTidakTarget = 0;
        let validAgentCount = 0;
        let sumAht = 0;

        groupRows.forEach((row, idx) => {
            const r = idx + 2;
            const dataRow = ws.getRow(r);

            // Validate if row has real data (ID Oceana must vary)
            const idVal = row[map.id];

            dataRow.getCell(1).value = idx + 1;
            dataRow.getCell(2).value = idVal;
            dataRow.getCell(3).value = row[map.name];
            dataRow.getCell(4).value = row[map.skill];

            const ahtVal = Number(row[map.aht]) || 0;
            dataRow.getCell(5).value = ahtVal;

            let targetAHT = 0;
            const skillKey = row[map.skill] ? row[map.skill].toString().toUpperCase() : "";
            if (TARGET_AHT_MAP[skillKey]) targetAHT = TARGET_AHT_MAP[skillKey];
            else if (TARGET_AHT_MAP[skillKey.replace(/\s/g, '')]) targetAHT = TARGET_AHT_MAP[skillKey.replace(/\s/g, '')];
            dataRow.getCell(6).value = targetAHT;

            // Logic 'Sesuai AHT'
            // Using formula in Excel so user can see it live, but we also track count for summary
            const isTarget = ahtVal <= targetAHT;

            if (idVal) { // Only count logic if valid agnet
                if (isTarget) totalTarget++;
                else totalTidakTarget++;
                validAgentCount++;
                sumAht += ahtVal;
            }

            // We use simple string for safety instead of formula if preferred, 
            // but user asked for automation. Let's stick to Formula for cell G.
            dataRow.getCell(7).value = { formula: `IF(E${r}<=F${r},"Target","Tidak Target")` };

            // Styling logic specific columns ? (Maybe later)

            let valPct = map.pct !== -1 ? row[map.pct] : null;
            if (typeof valPct === 'number' && valPct > 1) valPct = valPct / 100;
            else if (typeof valPct === 'string') {
                let parsed = parseFloat(valPct.replace('%', ''));
                if (!isNaN(parsed)) {
                    if (parsed > 1) parsed = parsed / 100;
                    valPct = parsed;
                }
            }
            dataRow.getCell(8).value = valPct;
            dataRow.getCell(8).numFmt = '0.00%';

            // Border for data row
            for (let c = 1; c <= 8; c++) {
                dataRow.getCell(c).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
            }
        });

        const lastDataRow = 1 + groupRows.length; // Row A2 to A[lastDataRow]

        // ---------------------------------------------------------
        // ADD SUMMARY AS PER USER IMAGE (USING FORMULAS)
        // ---------------------------------------------------------

        // 1. Total Target (Below Col G)
        const summaryRow1 = ws.getRow(lastDataRow + 1);

        const cellTotalTgtLabel = summaryRow1.getCell(6); // F
        cellTotalTgtLabel.value = "Total Target:";
        cellTotalTgtLabel.font = { italic: true };
        cellTotalTgtLabel.alignment = { horizontal: 'right' };

        const cellTotalTgtVal = summaryRow1.getCell(7); // G
        /* Formula: =COUNTIF(G2:G[lastDataRow], "Target") */
        cellTotalTgtVal.value = { formula: `COUNTIF(G2:G${lastDataRow},"Target")` };
        cellTotalTgtVal.font = { bold: true };
        cellTotalTgtVal.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        // 2. Total Tidak Target (Below Total Target)
        const summaryRow2 = ws.getRow(lastDataRow + 2);
        const cellTotalTidakLabel = summaryRow2.getCell(6); // F
        cellTotalTidakLabel.value = "Total Tidak Target:";
        cellTotalTidakLabel.font = { italic: true };
        cellTotalTidakLabel.alignment = { horizontal: 'right' };

        const cellTotalTidakVal = summaryRow2.getCell(7); // G
        /* Formula: =COUNTIF(G2:G[lastDataRow], "Tidak Target") */
        cellTotalTidakVal.value = { formula: `COUNTIF(G2:G${lastDataRow},"Tidak Target")` };
        cellTotalTidakVal.font = { bold: true };
        cellTotalTidakVal.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        // 3. Average AHT (Below Col E)
        const cellAvgAht = summaryRow1.getCell(5); // E
        /* Formula: ROUND(AVERAGE(...), 0) to remove decimals */
        cellAvgAht.value = { formula: `ROUND(AVERAGE(E2:E${lastDataRow}),0)` };
        cellAvgAht.font = { bold: true };
        cellAvgAht.numFmt = '0'; // Integer format
        cellAvgAht.alignment = { horizontal: 'center' }; // Center align specific
        cellAvgAht.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };


        // ---------------------------------------------------------
        // 2. DETAILED SUMMARY BOX (Starting a few rows below)
        // ---------------------------------------------------------
        const startSumRow = lastDataRow + 5;
        const totalTargetCellRef = `G${lastDataRow + 1}`; // Reference cell containing Total Target
        const totalTidakTargetCellRef = `G${lastDataRow + 2}`; // Reference cell containing Total Tidak Target
        const avgAhtCellRef = `E${lastDataRow + 1}`; // Reference cell containing Avg AHT

        // ... (Total Agent, Agent Target, % Target Agent, % Ready blocks remain same) ...
        // Keeping previous logic for lines 376-406 intact by starting replacement later or just re-pasting?
        // Re-pasting entire block for safety to ensure context match.

        // --- Total Agent ---
        ws.getCell(`C${startSumRow}`).value = "Total Agent";
        ws.getCell(`C${startSumRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        ws.getCell(`D${startSumRow}`).value = { formula: `${totalTargetCellRef}+${totalTidakTargetCellRef}` };
        ws.getCell(`D${startSumRow}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        // --- Agent Target ---
        ws.getCell(`C${startSumRow + 1}`).value = "Agent Target";
        ws.getCell(`C${startSumRow + 1}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        ws.getCell(`D${startSumRow + 1}`).value = { formula: totalTargetCellRef };
        ws.getCell(`D${startSumRow + 1}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        // --- % Target Agent ---
        ws.getCell(`C${startSumRow + 2}`).value = "% Target Agent";
        ws.getCell(`C${startSumRow + 2}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        ws.getCell(`D${startSumRow + 2}`).value = { formula: `D${startSumRow + 1}/D${startSumRow}` };
        ws.getCell(`D${startSumRow + 2}`).numFmt = '0.00%';
        ws.getCell(`D${startSumRow + 2}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        // --- Spacer ---

        // --- % Ready ---
        ws.getCell(`C${startSumRow + 4}`).value = "% Ready";
        ws.getCell(`C${startSumRow + 4}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        ws.getCell(`D${startSumRow + 4}`).value = { formula: `AVERAGE(H2:H${lastDataRow})` };
        ws.getCell(`D${startSumRow + 4}`).numFmt = '0.00%';
        ws.getCell(`D${startSumRow + 4}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };


        // --- Rata Rata Handling ---
        ws.getCell(`C${startSumRow + 6}`).value = "Rata Rata Handling";
        ws.getCell(`C${startSumRow + 6}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        /* Value: simply the number of seconds (Integer) */
        ws.getCell(`D${startSumRow + 6}`).value = { formula: avgAhtCellRef };
        ws.getCell(`D${startSumRow + 6}`).numFmt = '0'; // Integer (Detik)
        ws.getCell(`D${startSumRow + 6}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        // --- Menit / Detik ---
        ws.getCell(`C${startSumRow + 7}`).value = "Menit / Detik";
        ws.getCell(`C${startSumRow + 7}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        /* Value: 60 (Constant for reference?) Or divider. User asked "Untuk Menit / Detik (Detik)" -> Value 60. */
        ws.getCell(`D${startSumRow + 7}`).value = 60;
        ws.getCell(`D${startSumRow + 7}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        // --- Menit (Converted) ---
        ws.getCell(`C${startSumRow + 8}`).value = "Menit";
        ws.getCell(`C${startSumRow + 8}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        /* 
           Col D: Convert Seconds (Row-2) to Time (Division by 86400)
           Format: "mm:ss" as requested.
        */
        ws.getCell(`D${startSumRow + 8}`).value = { formula: `D${startSumRow + 6}/86400` };
        ws.getCell(`D${startSumRow + 8}`).numFmt = 'mm:ss';
        ws.getCell(`D${startSumRow + 8}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        /* 
           Col E: Same as D
        */
        ws.getCell(`E${startSumRow + 8}`).value = { formula: `D${startSumRow + 8}` };
        ws.getCell(`E${startSumRow + 8}`).numFmt = 'mm:ss';
        ws.getCell(`E${startSumRow + 8}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

        // ---------------------------------------------------------
        // 3. AUTO-FIT COLUMNS + ALIGNMENT FIX
        // ---------------------------------------------------------
        ws.columns.forEach((column, colIdx) => {
            let maxLength = 0;
            const colNum = colIdx + 1; // 1-based

            // Iterate over all cells in column
            column.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
                let cellLength = 0;
                if (cell.value) {
                    // Approximate text length
                    if (typeof cell.value === 'object' && cell.value.formula) {
                        cellLength = 10;
                    } else {
                        cellLength = cell.value.toString().length;
                    }
                }
                if (cellLength > maxLength) maxLength = cellLength;
            });
            // Set width
            column.width = maxLength < 10 ? 10 : maxLength + 2;

            // CENTER ALIGNMENT LOGIC (Point 2)
            // Columns: No(1), ID(2), Skill(4), AHT(5), Target(6), Sesuai(7), Ready(8)
            // Agent Name(3) excluded
            if ([1, 2, 4, 5, 6, 7, 8].includes(colNum)) {
                // Apply center alignment to Header (Row 1)
                const headerCell = ws.getRow(1).getCell(colNum);
                headerCell.alignment = { horizontal: 'center', vertical: 'middle' };

                // Apply check to Data Rows
                // Data is from Row 2 to lastDataRow
                for (let r = 2; r <= lastDataRow; r++) {
                    const cell = ws.getRow(r).getCell(colNum);
                    // Preserve borders if any, just update alignment
                    // Creating new object to avoid reference issues
                    const existingBorder = cell.border || {};
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    cell.border = existingBorder;
                }
            }
        });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    
    let fileName = 'Result.xlsx';
    if (originalFileName) {
        if (originalFileName.toLowerCase().endsWith('.xlsx')) {
            fileName = originalFileName.replace(/\.xlsx$/i, '_Result.xlsx');
        } else if (originalFileName.toLowerCase().endsWith('.xls')) {
            fileName = originalFileName.replace(/\.xls$/i, '_Result.xlsx');
        } else if (originalFileName.toLowerCase().endsWith('.csv')) {
            fileName = originalFileName.replace(/\.csv$/i, '_Result.xlsx');
        } else {
            fileName = originalFileName + '_Result.xlsx';
        }
    }
    
    saveAs(new Blob([buffer]), fileName);
};


// ---------------------------------------------------------------------------
// 3. GENERATE CMS REPORT (UPDATED WITH 6 COLUMN SUPPORT)
// ---------------------------------------------------------------------------

const readWorkbook = async (file) => {
    const data = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data);
    return wb;
};

export const generateCMSReport = async (files, targetMonth) => {
    console.log("Starting CMS Report Generation...");

    if (!files || files.length === 0) {
        alert("No files provided!");
        return;
    }
    if (!targetMonth) {
        alert("Please select a month!");
        return;
    }

    const summaryFile = files.find(f => f.name.toLowerCase().includes('summary'));
    const dailyFilesRaw = files.filter(f => f !== summaryFile);

    if (!summaryFile) {
        alert("Summary.xlsx not found. Please upload it.");
        return;
    }

    // 1. Sort Daily Files Strictly
    const dailyFiles = dailyFilesRaw
        .map(f => ({ file: f, day: getDayFromFileName(f.name) }))
        .filter(item => item.day !== null)
        .sort((a, b) => a.day - b.day)
        .map(item => item.file);

    try {
        const summaryWb = await readWorkbook(summaryFile);
        const summarySheet = summaryWb.worksheets.find(ws => ws.name.trim().toLowerCase() === targetMonth.toLowerCase());

        if (!summarySheet) {
            alert(`Sheet "${targetMonth}" not found.`);
            return;
        }

        // -------------------------------------------------------------------
        // 2. LOCATE DESTINATION DATE HEADERS
        // -------------------------------------------------------------------
        let dateBlockMap = {};
        let intervalDateRowIdx = -1;

        console.log("Scanning for Date Headers in Summary...");
        for (let r = 20; r <= 40; r++) {
            const row = summarySheet.getRow(r);
            const tempMap = {};
            let count = 0;

            row.eachCell((cell, colNumber) => {
                let val = cell.value;
                // Handle Rich Text / Objects in Header
                if (val && typeof val === 'object') {
                    if (val.result) val = val.result;
                    else if (val.richText && Array.isArray(val.richText)) {
                        val = val.richText.map(rt => rt.text).join('');
                    } else if (val.text) {
                        val = val.text;
                    }
                }

                if (!val) return;

                let day = null;
                if (val instanceof Date) {
                    day = val.getDate();
                } else if (typeof val === 'string') {
                    // Match "01 Jan", "1-Jan", "Jan 1"
                    const match = val.trim().match(/^(\d{1,2})[\s\-\/]/);
                    if (match) day = parseInt(match[1], 10);
                }

                if (day && day >= 1 && day <= 31) {
                    tempMap[day] = colNumber;
                    count++;
                }
            });

            if (count >= 1) { // Even 1 date found is better than nothing
                // Merge into main map (sometimes dates are on different rows?)
                // Usually they are on one row.
                if (count > Object.keys(dateBlockMap).length) {
                    dateBlockMap = tempMap;
                    intervalDateRowIdx = r;
                }
                console.log(`Potential Date Header Row ${r}. Dates found: ${count}`);
            }
        }

        // Fallback: If Day 1 (01) is not found, User says "Paste di B30" for Day 1.
        // So Day 1 -> Column 2.
        if (!dateBlockMap[1]) {
            console.warn("Day 1 Header not detected. Defaulting to Column 2 (B) as per user request.");
            dateBlockMap[1] = 2; // Column B
        }

        if (Object.keys(dateBlockMap).length === 0) {
            alert("Could not detect Date headers!");
            return;
        }

        // -------------------------------------------------------------------
        // 2b. FIXED TARGET COLUMN MAPPING (User Specified)
        // -------------------------------------------------------------------
        // Mapping relative to the Date Start Column (e.g., if Date is at B, that is Offset 0):
        // 1. Offered (B)      -> Offset 0
        // 2. Ans Within 20s (C)-> Offset 1
        // 3. Answered (D)     -> Offset 2
        // 4. Abandoned (E)    -> Offset 3
        // 5. ABN Calls > 5s (F)-> Offset 4
        // 6. % SL Within 20s (G)-> Offset 5

        const targetOffsets = {
            offered: 0,
            acceptable: 1,
            acd: 2,
            abn: 3,
            abn_gt_5: 4,
            sl: 5
        };
        console.log("Using Fixed Target Offsets:", targetOffsets);

        // -------------------------------------------------------------------
        // 3. PROCESS FILES
        // -------------------------------------------------------------------
        // -------------------------------------------------------------------
        // 3. PROCESS FILES WITH FORCED COLUMN CALCULATION
        // -------------------------------------------------------------------
        for (const file of dailyFiles) {
            const dayNum = getDayFromFileName(file.name);
            if (!dayNum) continue;

            // FORCE CALCULATION: Day 1 -> Col 2 (B). Width 6.
            // Formula: StartCol = 2 + (dayNum - 1) * 6
            const startCol = 2 + (dayNum - 1) * 6;

            const dailyWb = await readWorkbook(file);
            const dailySheet = dailyWb.worksheets[0];

            // -------------------------------------------------------
            // 3b. COPY DAILY AVERAGES (Top Summary)
            // -------------------------------------------------------
            const summaryColIdx = dayNum + 1; // 1->2(B), 2->3(C)...

            // Helper to parse Time String "00:04:35" -> Excel Serial Number
            const parseTime = (val) => {
                if (!val) return 0;
                if (typeof val === 'number') return val; // Already serial
                if (val instanceof Date) {
                    // Convert JS Date to Excel Serial (roughly)
                    // Or just return the Date object, ExcelJS handles it.
                    return val;
                }
                if (typeof val === 'string') {
                    // Try parsing "HH:mm:ss"
                    const parts = val.trim().split(':');
                    if (parts.length === 3) {
                        const h = parseInt(parts[0], 10) || 0;
                        const m = parseInt(parts[1], 10) || 0;
                        const s = parseInt(parts[2], 10) || 0;
                        // Excel Time = (H*3600 + M*60 + S) / 86400
                        return (h * 3600 + m * 60 + s) / 86400;
                    }
                }
                return 0; // Fallback
            };

            const dailyRow2 = dailySheet.getRow(2);

            // Parse strictly to ensure they are Numbers/Serials for the AVERAGE formula to work
            const valActiveTime = parseTime(getSafeVal(dailyRow2, 10));
            const valHoldTime = parseTime(getSafeVal(dailyRow2, 11));

            // Write to Summary with Number Format
            const cellActive = summarySheet.getCell(13, summaryColIdx);
            cellActive.value = valActiveTime;
            cellActive.numFmt = '[h]:mm:ss'; // Ensure it looks like time

            const cellHold = summarySheet.getCell(14, summaryColIdx);
            cellHold.value = valHoldTime;
            cellHold.numFmt = '[h]:mm:ss';

            // ... (srcCols definition remains below)

            const srcCols = {
                offered: 4,     // D
                acceptable: 5,  // E
                acd: 6,         // F
                abn: 7,         // G
                abn_gt_5: 8,    // H
                sl: 9           // I
            };

            const DATA_START_ROW = 3;
            // User requested starts at Row 30
            const TARGET_ROW_START = 30;

            let writeCount = 0;

            dailySheet.eachRow((row, rowNumber) => {
                // Skip Header
                if (rowNumber < DATA_START_ROW) return;

                // Safety Cap: Prevent processing thousands of empty rows (e.g. infinite scrolling sheets)
                if (rowNumber > 1000) return;

                // Skip Completely Empty Rows to avoid corrupting Summary XML with thousands of empty tag updates
                const hasData = [srcCols.offered, srcCols.acceptable, srcCols.acd].some(idx => {
                    const v = row.getCell(idx).value;
                    return v !== null && v !== undefined && v !== '';
                });
                if (!hasData) return;

                const targetRow = summarySheet.getRow(TARGET_ROW_START + writeCount);
                writeCount++;

                // Strict Mapping:
                // Offered (Source D) -> Target Start + 0
                targetRow.getCell(startCol + 0).value = getSafeVal(row, srcCols.offered);
                // Ans Within 20s (Source E) -> Target Start + 1
                targetRow.getCell(startCol + 1).value = getSafeVal(row, srcCols.acceptable);
                // Answered (Source F) -> Target Start + 2
                targetRow.getCell(startCol + 2).value = getSafeVal(row, srcCols.acd);
                // Abandoned (Source G) -> Target Start + 3
                targetRow.getCell(startCol + 3).value = getSafeVal(row, srcCols.abn);
                // ABN Calls > 5s (Source H) -> Target Start + 4
                targetRow.getCell(startCol + 4).value = getSafeVal(row, srcCols.abn_gt_5);
                // Service Level (Source I) -> Target Start + 5
                targetRow.getCell(startCol + 5).value = getSafeVal(row, srcCols.sl);
            });

            // Yield after each file to keep UI responsive
            await new Promise(resolve => setTimeout(resolve, 0));
        }

        // -------------------------------------------------------------------
        // 4. ADD TOTAL/AVERAGE FORMULAS (Row 13 & 14, Col AG)
        // -------------------------------------------------------------------
        // Calculate Average for the month (Columns B to AF)
        summarySheet.getCell('AG13').value = { formula: 'AVERAGE(B13:AF13)' };
        summarySheet.getCell('AG14').value = { formula: 'AVERAGE(B14:AF14)' };

        const buffer = await summaryWb.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), 'Summary_Updated.xlsx');
        console.log("Report Generation Complete.");

    } catch (err) {
        console.error("CMS Report Error:", err);
        alert(`Error: ${err.message}`);
    }
};
// ---------------------------------------------------------------------------
// 4. JOIN SHEETS (Large File Robust Version - Async to prevent UI freezing)
// ---------------------------------------------------------------------------
export const joinSheets = async (workbook, onProgress) => {
    // workbook is already parsed by XLSX.read in App.jsx
    let allData = [];
    let masterHeaders = [];
    let sheetCount = 0;

    const clean = (s) => s ? s.toString().toLowerCase().trim().replace(/\s+/g, ' ') : '';

    console.log(`--- STARTING JOIN PROCESS ---`);
    console.log(`Total sheets found in workbook: ${workbook.SheetNames.length}`);

    // Change to async loop to prevent UI freezing
    for (let sIdx = 0; sIdx < workbook.SheetNames.length; sIdx++) {
        const sheetName = workbook.SheetNames[sIdx];
        const worksheet = workbook.Sheets[sheetName];
        
        // Yield to UI thread every sheet to keep animations spinning
        await new Promise(resolve => setTimeout(resolve, 0));

        if (!worksheet) continue;

        // Convert to AOA with forced range 0 to bypass bad metadata
        const rows = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1, 
            blankrows: false,
            defval: "",
            range: 0 
        });

        if (!rows || rows.length === 0) {
            console.log(`Sheet "${sheetName}" has no data. Skipping.`);
            continue;
        }

        let bestHeaderIdx = 0; // Default to first row

        if (sheetCount === 0) {
            // Master Sheet: Find the row with the most content in the first 40 rows
            let maxCols = -1;
            for (let i = 0; i < Math.min(rows.length, 40); i++) {
                const count = rows[i].filter(c => c !== "" && c !== null).length;
                if (count > maxCols) {
                    maxCols = count;
                    bestHeaderIdx = i;
                }
            }
            
            masterHeaders = rows[bestHeaderIdx];
            const dataRows = rows.slice(bestHeaderIdx + 1);
            
            // Add "Source Sheet" as first column
            const finalHeaders = ["Source Sheet", ...masterHeaders];
            const processedRows = dataRows.map(r => [sheetName, ...r]);
            
            allData = [finalHeaders, ...processedRows];
            sheetCount++;
            console.log(`Sheet 1 ("${sheetName}"): Header at row ${bestHeaderIdx + 1}. Rows: ${processedRows.length}`);
        } else {
            // Subsequent sheets: Fuzzy match master headers
            let bestMatchCount = -1;
            const masterClean = masterHeaders.map(clean);

            for (let i = 0; i < Math.min(rows.length, 40); i++) {
                const currentClean = rows[i].map(clean);
                let matches = 0;
                currentClean.forEach(c => {
                    if (c !== "" && masterClean.includes(c)) matches++;
                });

                if (matches > bestMatchCount) {
                    bestMatchCount = matches;
                    bestHeaderIdx = i;
                }
            }

            const currentHeaders = rows[bestHeaderIdx];
            const dataRows = rows.slice(bestHeaderIdx + 1);
            
            // Create mapping: current index -> master index
            const currentToMaster = {};
            const currentClean = currentHeaders.map(clean);
            currentClean.forEach((c, idx) => {
                if (c === "") return;
                const mIdx = masterClean.indexOf(c);
                if (mIdx !== -1) currentToMaster[idx] = mIdx;
            });

            let addedInThisSheet = 0;
            const rowsToAdd = dataRows.map(row => {
                // Skip completely empty rows
                if (!row || row.every(c => c === "" || c === null)) return null;

                const alignedRow = new Array(masterHeaders.length + 1).fill("");
                alignedRow[0] = sheetName;

                let hasData = false;
                
                // If fuzzy mapping found some columns, use them. 
                // Otherwise, fall back to direct index mapping (assuming same structure).
                if (Object.keys(currentToMaster).length > 0) {
                    row.forEach((val, colIdx) => {
                        if (currentToMaster[colIdx] !== undefined) {
                            alignedRow[currentToMaster[colIdx] + 1] = val;
                            if (val !== "" && val !== null) hasData = true;
                        }
                    });
                } else {
                    // Fall-back: Direct index mapping
                    row.forEach((val, colIdx) => {
                        if (colIdx < masterHeaders.length) {
                            alignedRow[colIdx + 1] = val;
                            if (val !== "" && val !== null) hasData = true;
                        }
                    });
                }

                if (hasData) {
                    addedInThisSheet++;
                    return alignedRow;
                }
                return null;
            }).filter(r => r !== null);

            // Safer push for large arrays instead of spread operator
            for (let i = 0; i < rowsToAdd.length; i++) {
                allData.push(rowsToAdd[i]);
            }
            sheetCount++;
            console.log(`Sheet ${sheetCount} ("${sheetName}"): Added ${addedInThisSheet} rows.`);
        }
        
        // Update progress based on sheets processed
        if (onProgress) {
            const pct = Math.round(((sIdx + 1) / workbook.SheetNames.length) * 100);
            onProgress(pct);
        }
    }


    console.log(`Join complete. Total Rows: ${allData.length}`);
    return { data: allData, sheetCount };
};
