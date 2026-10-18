const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    icon: path.join(__dirname, '../public/Icons Excel 1.png'),
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      webSecurity: false,
    },
  });

  const startUrl = process.env.ELECTRON_START_URL || 'http://localhost:5173';

  if (process.env.NODE_ENV === 'development') {
    win.loadURL(startUrl);
  } else {
    win.loadFile(path.join(__dirname, '../dist/index.html'));
  }
}

app.commandLine.appendSwitch('js-flags', '--max-old-space-size=4096');

// IPC Handler for Save Dialog
ipcMain.handle('show-save-dialog', async (event, options) => {
  return await dialog.showSaveDialog(options);
});

// IPC Handler for Open Dialog (Native File Picker)
ipcMain.handle('open-file-dialog', async (event, options) => {
  return await dialog.showOpenDialog(options);
});

// IPC Handler for Background Joining & Streaming Export
ipcMain.handle('process-large-join', async (event, { inputPath, outputPath }) => {
  try {
    console.log(`Starting background join: ${inputPath} -> ${outputPath}`);
    
    // 1. Read Workbook with Dense Mode
    const workbook = XLSX.readFile(inputPath, { cellDates: true, dense: true });
    const sheetNames = workbook.SheetNames;
    
    // 2. Initialize ExcelJS Streaming Writer
    const wbWriter = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: outputPath,
      useStyles: true,
      useSharedStrings: false // Essential for memory efficiency
    });
    
    const ws = wbWriter.addWorksheet('Joined Data');
    
    // Header Style
    const headerStyle = {
      font: { bold: true, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } },
      alignment: { vertical: 'middle', horizontal: 'center' },
      border: {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      }
    };

    let masterHeaders = null;
    let sheetCount = 0;

    for (let sIdx = 0; sIdx < sheetNames.length; sIdx++) {
      const sheetName = sheetNames[sIdx];
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 0, defval: "" });
      
      if (!rows || rows.length === 0) continue;

      if (sheetCount === 0) {
        // First sheet: find header
        let bestHeaderIdx = 0;
        let maxCols = -1;
        for (let i = 0; i < Math.min(rows.length, 20); i++) {
           const count = rows[i].filter(c => c !== null && String(c).trim() !== "").length;
           if (count > maxCols) { maxCols = count; bestHeaderIdx = i; }
        }
        masterHeaders = rows[bestHeaderIdx];
        
        // Write Master Header
        const headerRow = ws.getRow(1);
        headerRow.values = ['Source Sheet', ...masterHeaders];
        headerRow.eachCell(cell => { cell.style = headerStyle; });
        headerRow.commit();
        sheetCount++;
        
        // Process data
        for (let i = bestHeaderIdx + 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0 || row.every(c => c === null || String(c).trim() === "")) continue;
          ws.addRow([sheetName, ...row]).commit();
        }
      } else {
        // Subsequent sheets
        const masterHeaderStr = masterHeaders ? masterHeaders.map(c => String(c || '').trim().toLowerCase()).join('|') : "";
        
        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          
          // 1. Skip completely empty rows
          const isRowEmpty = !row || row.length === 0 || row.every(c => c === null || String(c).trim() === "");
          if (isRowEmpty) continue;
          
          // 2. Precise Header Check
          const currentRowStr = row.map(c => String(c || '').trim().toLowerCase()).join('|');
          
          // ONLY skip if it matches the master header exactly (ignoring case/spaces)
          // OR if the first column is exactly "no" and it's within the first few rows
          if (currentRowStr === masterHeaderStr) continue;
          
          if (i < 5) {
             const firstCol = String(row[0] || '').trim().toLowerCase();
             if (firstCol === 'no') continue;
          }

          ws.addRow([sheetName, ...row]).commit();
        }
      }
    }

    await ws.commit();
    await wbWriter.commit();
    console.log('Background join complete.');
    return { success: true };
  } catch (err) {
    console.error('Background join error:', err);
    return { success: false, error: err.message };
  }
});

// IPC Handler to get stats without loading data into Renderer
ipcMain.handle('get-join-stats', async (event, inputPath) => {
  try {
    const workbook = XLSX.readFile(inputPath, { cellDates: true, dense: true });
    const sheetNames = workbook.SheetNames;
    let totalRows = 0;
    
    for (const name of sheetNames) {
      const sheet = workbook.Sheets[name];
      if (!sheet['!ref']) continue;
      const range = XLSX.utils.decode_range(sheet['!ref']);
      totalRows += (range.e.r - range.s.r); // Approximate row count
    }
    
    return { success: true, sheetCount: sheetNames.length, totalRows };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
