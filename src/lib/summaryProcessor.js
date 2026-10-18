import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

export async function processSummaryReport(originalData, files) {
    if (!originalData || originalData.length === 0) {
        throw new Error("Tidak ada data untuk diproses.");
    }

    try {
        let fs;
        if (typeof window.require !== 'undefined') {
            try {
                fs = window.require('fs');
            } catch (e) {
                console.warn("Node fs not available");
            }
        }

        let fileBuffer;
        
        // Coba cari file summary dari daftar file yang diunggah
        if (files && files.length > 0) {
            for (let i = 0; i < files.length; i++) {
                if (files[i].name && files[i].name.toLowerCase().includes('summary')) {
                    fileBuffer = await files[i].arrayBuffer();
                    break;
                }
            }
        }

        // Jika tidak ada file summary yang diunggah dan ada fs, coba baca dari path lokal
        if (!fileBuffer && fs) {
            const summaryPath = 'E:\\Aplikasi Lokal\\Otomasi Report\\Summary Sasaran Mutu All Agent Montly.xlsx';
            try {
                fileBuffer = fs.readFileSync(summaryPath);
            } catch (err) {
                // fall through to error message below
            }
        }

        if (!fileBuffer) {
            throw new Error("Gagal membaca file Summary. Di browser/live demo, pastikan Anda MENGUNGGAH file 'Summary Sasaran Mutu All Agent Montly.xlsx' secara bersamaan dengan file data.");
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(fileBuffer);
        const sheet = workbook.worksheets[0]; // "Sheet (A1)"

        if (!sheet) {
            throw new Error("Sheet pertama tidak ditemukan di file Summary.");
        }

        const norm = (h) => (h ? h.toString().toLowerCase().replace(/[^a-z0-9]/g, '') : '');
        const inputHeaders = originalData[0];
        const map = {
            id: inputHeaders.findIndex(h => norm(h).includes('idoceana')),
            skill: inputHeaders.findIndex(h => norm(h) === 'skill'),
            aht: inputHeaders.findIndex(h => norm(h).includes('ahtbyacd')),
            pct: inputHeaders.findIndex(h => norm(h).includes('ready') || norm(h).includes('%ready'))
        };
        if (map.skill === -1) map.skill = 3;
        if (map.id === -1) map.id = 1;
        if (map.aht === -1) map.aht = 4;

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

        const groups = {};
        const rows = originalData.slice(1);
        rows.forEach(row => {
            const skill = row[map.skill] ? row[map.skill].toString().trim().toUpperCase() : 'UNCATEGORIZED';
            if (!groups[skill]) groups[skill] = [];
            groups[skill].push(row);
        });

        const groupAggregates = {};
        for (const skill of Object.keys(groups)) {
            const groupRows = groups[skill];
            let totalTarget = 0;
            let totalTidakTarget = 0;
            let sumAht = 0;
            let sumPct = 0;
            let countPct = 0;
            let validAgentCount = 0;

            groupRows.forEach(row => {
                const idVal = row[map.id];
                const ahtVal = Number(row[map.aht]) || 0;
                const skillKey = skill.toUpperCase();
                let targetAHT = 0;
                if (TARGET_AHT_MAP[skillKey]) {
                    targetAHT = TARGET_AHT_MAP[skillKey];
                } else if (TARGET_AHT_MAP[skillKey.replace(/\s/g, '')]) {
                    targetAHT = TARGET_AHT_MAP[skillKey.replace(/\s/g, '')];
                }

                const isTarget = ahtVal <= targetAHT;
                if (idVal) {
                    if (isTarget) totalTarget++;
                    else totalTidakTarget++;
                    validAgentCount++;
                    sumAht += ahtVal;
                }

                let valPct = map.pct !== -1 ? row[map.pct] : null;
                if (typeof valPct === 'number' && valPct > 1) valPct = valPct / 100;
                else if (typeof valPct === 'string') {
                    let parsed = parseFloat(valPct.replace('%', ''));
                    if (!isNaN(parsed)) {
                        if (parsed > 1) parsed = parsed / 100;
                        valPct = parsed;
                    }
                }
                if (valPct != null) {
                    sumPct += valPct;
                    countPct++;
                }
            });

            const totalAgent = totalTarget + totalTidakTarget;
            const agentTarget = totalTarget;
            const avgAht = validAgentCount > 0 ? sumAht / validAgentCount : 0;
            const avgPct = countPct > 0 ? sumPct / countPct : 0;
            const avgAhtMinutes = Math.round(avgAht) / 86400;

            groupAggregates[skill] = {
                totalAgent,
                agentTarget,
                rawAvgAht: avgAht, // Added raw average in seconds
                avgPct
            };
        }

        for (let row = 2; row <= 11; row++) {
            const sheetNameCell = sheet.getCell(`A${row}`).value;
            if (!sheetNameCell) continue;
            
            const sheetNameStr = String(sheetNameCell).trim().toUpperCase();
            const agg = groupAggregates[sheetNameStr];
            
            if (agg) {
                const totalAgent = agg.totalAgent;
                const agentTarget = agg.agentTarget;
                const pctTargetAgent = totalAgent > 0 ? (agentTarget / totalAgent) * 100 : 0;
                
                const rawAvgAht = Math.round(agg.rawAvgAht || 0);
                const mm = Math.floor(rawAvgAht / 60);
                const ss = (rawAvgAht % 60).toString().padStart(2, '0');
                
                const combinedAHT = `${mm}:${ss} / ${pctTargetAgent.toFixed(2)}%`;
                
                const ahtCell = sheet.getCell(`D${row}`);
                ahtCell.value = combinedAHT;
                ahtCell.numFmt = '@'; // Force text format

                const readyCell = sheet.getCell(`F${row}`);
                readyCell.value = agg.avgPct;
                readyCell.numFmt = '0.00%';
            }
        }

        const agentMap = {
            'bk': { row: 16 },
            'kk': { row: 20 },
            'bpc': { row: 24 },
            'emerald': { row: 28 },
            'bs': { row: 32, row2: 52 },
            't&r': { row: 36 },
            'mc': { row: 40, row2: 48 },
            'sat': { row: 44 }
        };

        for (const [key, mapping] of Object.entries(agentMap)) {
            const agg = groupAggregates[key.toUpperCase()];
            if (agg) {
                const totalAgent = agg.totalAgent;
                const agentTarget = agg.agentTarget;

                sheet.getCell(`C${mapping.row}`).value = totalAgent;
                sheet.getCell(`C${mapping.row + 1}`).value = agentTarget;
                const formulaCell = sheet.getCell(`C${mapping.row + 2}`);
                formulaCell.value = { formula: `C${mapping.row + 1}/C${mapping.row}`, result: totalAgent ? agentTarget/totalAgent : 0 };
                formulaCell.numFmt = '0.00%';

                if (mapping.row2) {
                    sheet.getCell(`C${mapping.row2}`).value = totalAgent;
                    sheet.getCell(`C${mapping.row2 + 1}`).value = agentTarget;
                    const formulaCell2 = sheet.getCell(`C${mapping.row2 + 2}`);
                    formulaCell2.value = { formula: `C${mapping.row2 + 1}/C${mapping.row2}`, result: totalAgent ? agentTarget/totalAgent : 0 };
                    formulaCell2.numFmt = '0.00%';
                }
            }
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        return blob;

    } catch (error) {
        console.error(error);
        throw error;
    }
}
