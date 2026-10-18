import React, { useState, useEffect, useRef } from 'react';
import { FileUpload } from './components/FileUpload';
import { DataTable } from './components/DataTable';
import { motion, AnimatePresence } from 'framer-motion';
import { LayoutDashboard, Settings, Github, Database, Wand2, FileSpreadsheet, BarChart3, Clock, Users, Activity, Search, Download, Filter, AlertTriangle, Plus, Trash2 } from 'lucide-react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { generateAHTReport, generateCMSReport, generatePreviewData, joinSheets } from './lib/generateReport';
import { saveAs } from 'file-saver';
import { processSummaryReport } from './lib/summaryProcessor';


const INITIAL_REPORT_MENUS = [
  { id: 'report-sasaran', label: 'Report Sasaran Mutu', icon: FileSpreadsheet },
  { id: 'report-daily', label: 'Report CMS', icon: Activity },
  { id: 'report-performance', label: 'Report Performance Agent', icon: BarChart3 },
  { id: 'join-sheet', label: 'Join Sheet', icon: FileSpreadsheet },
];

function App() {
  // State
  const [dataState, setDataState] = useState({
    fileName: null,
    data: null,
    originalData: null // Store original for final generation
  });
  const [isGenerating, setIsGenerating] = useState(false);
  const [isPreviewMode, setIsPreviewMode] = useState(false);
  const [progress, setProgress] = useState(0);
  const [previewSheet, setPreviewSheet] = useState('All');
  const [errorInfo, setErrorInfo] = useState(null); // { title: string, message: string }

  const [searchTerm, setSearchTerm] = useState('');

  // CMS Month Selection State
  const [showMonthSelector, setShowMonthSelector] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState('Jan');

  const MONTH_OPTIONS = [
    "Jan", "Feb", "Mar", "Apr", "Mei", "Juni",
    "Juli", "Agust", "Sept", "Okto", "Nov", "Des"
  ];

  // Navigation Confirmation State
  const [showNavConfirm, setShowNavConfirm] = useState(false);
  const [pendingTab, setPendingTab] = useState(null);

  // Navigation State
  const [activeTab, setActiveTab] = useState('report-sasaran');
  const [reportMenus, setReportMenus] = useState(() => {
    const saved = localStorage.getItem('custom_report_menus');
    const customMenus = [];
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        // Only take the ones that are actually custom (start with 'custom-')
        const filteredCustom = parsed.filter(m => m.id.startsWith('custom-')).map(menu => ({
          ...menu,
          icon: FileSpreadsheet
        }));
        customMenus.push(...filteredCustom);
      } catch (e) {
        console.error("Failed to load custom menus from storage", e);
      }
    }
    return [...INITIAL_REPORT_MENUS, ...customMenus];
  });

  useEffect(() => {
    // ONLY save custom menus to localStorage
    const customMenusToSave = reportMenus
      .filter(m => m.id.startsWith('custom-'))
      .map(({ id, label, password }) => ({ id, label, password }));
    localStorage.setItem('custom_report_menus', JSON.stringify(customMenusToSave));
  }, [reportMenus]);

  const [showAddModal, setShowAddModal] = useState(false);
  const [newMenuName, setNewMenuName] = useState('');
  const [newMenuPassword, setNewMenuPassword] = useState('');

  // Delete Menu State
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [menuToDelete, setMenuToDelete] = useState(null);
  const [deletePassword, setDeletePassword] = useState('');
  const [deleteError, setDeleteError] = useState(null);
  const [showSuccess, setShowSuccess] = useState(false);
  const [successMessage, setSuccessMessage] = useState('');
  const [failedAttempts, setFailedAttempts] = useState(0);
  const [currentTime, setCurrentTime] = useState(Date.now());
  const [lockoutUntil, setLockoutUntil] = useState(() => {
    const saved = localStorage.getItem('delete_lockout_until');
    if (saved) {
      const until = parseInt(saved);
      return until > Date.now() ? until : null;
    }
    return null;
  });
  const deleteInputRef = useRef(null);

  useEffect(() => {
    if (lockoutUntil) {
      localStorage.setItem('delete_lockout_until', lockoutUntil.toString());
      const timer = setInterval(() => {
        const now = Date.now();
        setCurrentTime(now);
        if (now >= lockoutUntil) {
          setLockoutUntil(null);
          setFailedAttempts(0);
          setDeleteError(null);
          localStorage.removeItem('delete_lockout_until');
        }
      }, 1000);
      return () => clearInterval(timer);
    }
  }, [lockoutUntil]);

  useEffect(() => {
    if (showDeleteModal) {
      const timer = setTimeout(() => {
        if (!lockoutUntil) deleteInputRef.current?.focus();
      }, 100);
      return () => clearTimeout(timer);
    } else {
      setDeleteError(null);
      setDeletePassword('');
    }
  }, [showDeleteModal, lockoutUntil]);
  const DELETE_PASSWORD = 'admin'; // Simple password for demo

  const handleDataLoaded = ({ fileName, data, files, multiple, rawFile }) => {
    let finalData = data;
    let finalOriginal = data;

    if (multiple && files && (!data || activeTab === 'join-sheet' || activeTab === 'report-daily')) {
      // Create a dummy preview for list of files only if no data was extracted or for specific tabs
      const fileList = [["File Name", "Size (KB)"]];
      files.forEach(f => {
        fileList.push([f.name, (f.size / 1024).toFixed(2)]);
      });
      finalData = fileList;
      finalOriginal = fileList;
    }

    setDataState({
      fileName,
      data: finalData,
      originalData: finalOriginal,
      files: files || null,
      rawFile: rawFile || null
    });
    setIsPreviewMode(false);
    setIsGenerating(false);
    setPreviewSheet('All');
    setSearchTerm('');
  };

  const handleReset = () => {
    setDataState({ fileName: null, data: null, originalData: null });
    setIsPreviewMode(false);
    setPreviewSheet('All');
    setSearchTerm('');
    setProgress(0);
  };

  const handleNavClick = (tabId) => {
    // If we are already on this tab, do nothing
    if (tabId === activeTab) return;

    // If there is data loaded, confirm before switching because it will be cleared
    if (dataState.originalData || dataState.data) {
      setPendingTab(tabId);
      setShowNavConfirm(true);
    } else {
      setActiveTab(tabId);
      handleReset();
    }
  };

  const confirmNavigation = () => {
    if (pendingTab) {
      setActiveTab(pendingTab);
      handleReset();
    }
    setShowNavConfirm(false);
    setPendingTab(null);
  };

  const cancelNavigation = () => {
    setShowNavConfirm(false);
    setPendingTab(null);
  };

  const handleAddReport = () => {
    if (!newMenuName.trim() || !newMenuPassword.trim()) return;
    
    const newId = `custom-${Date.now()}`;
    const newMenu = {
      id: newId,
      label: newMenuName,
      password: newMenuPassword,
      icon: FileSpreadsheet
    };
    
    setReportMenus(prev => [...prev, newMenu]);
    setNewMenuName('');
    setNewMenuPassword('');
    setShowAddModal(false);
    setActiveTab(newId);
    handleReset();
  };

  const handleDeleteMenu = () => {
    if (lockoutUntil && Date.now() < lockoutUntil) return;

    const targetPassword = menuToDelete?.password || 'admin';
    
    if (deletePassword !== targetPassword) {
      const newAttempts = failedAttempts + 1;
      setFailedAttempts(newAttempts);

      if (newAttempts >= 3) {
        const until = Date.now() + 15 * 60 * 1000;
        setLockoutUntil(until);
        setDeleteError("Password Salah 3X, Tunggu 15 menit untuk mencoba kembali");
      } else {
        setDeleteError(`Password salah! (${3 - newAttempts} kesempatan lagi)`);
      }
      
      setDeletePassword('');
      // Refocus after error
      setTimeout(() => {
        if (newAttempts < 3) deleteInputRef.current?.focus();
      }, 50);
      return;
    }

    if (menuToDelete) {
      const deletedLabel = menuToDelete.label;
      setReportMenus(prev => prev.filter(m => m.id !== menuToDelete.id));
      if (activeTab === menuToDelete.id) {
        setActiveTab('report-sasaran');
        handleReset();
      }
      setShowDeleteModal(false);
      setMenuToDelete(null);
      setDeletePassword('');
      setDeleteError(null);
      setFailedAttempts(0);
      
      // Show success popup
      setSuccessMessage(`Menu "${deletedLabel}" berhasil dihapus.`);
      setShowSuccess(true);
    }
  };

  const handleMonthConfirm = async () => {
    setShowMonthSelector(false);
    setIsGenerating(true);
    try {
      await generateCMSReport(dataState.files, selectedMonth);
    } catch (error) {
      console.error("Failed to generate CMS report:", error);
    } finally {
      setIsGenerating(false);
    }
  };

  const handleGenerateReport = async () => {
    // If no data, return
    if (!dataState.originalData) return;

    setIsGenerating(true);
    setProgress(0);
    try {
      if (activeTab === 'report-daily' && dataState.files) {
        // Show Month Selector instead of immediate generation
        setIsGenerating(false);
        setShowMonthSelector(true);
        return;
      } else if (activeTab === 'join-sheet') {
        if (!isPreviewMode) {
          // Wrap in timeout to allow UI to show loading state
          setTimeout(async () => {
            try {
              if (typeof window.require !== 'function') {
                throw new Error("Aplikasi tidak dapat menemukan modul Electron. Pastikan Anda menjalankan aplikasi ini di dalam Electron.");
              }
              
              const { ipcRenderer, webUtils } = window.require('electron');
              const rawFile = dataState.rawFile;
              let filePath = rawFile?.path;
              
              // Electron 28+ uses webUtils to get the path reliably
              if (!filePath && webUtils && rawFile) {
                try {
                  filePath = webUtils.getPathForFile(rawFile);
                  console.log("Path resolved via webUtils:", filePath);
                } catch (e) {
                  console.warn("webUtils.getPathForFile failed:", e);
                }
              }
              
              console.log("Raw File Object:", rawFile);
              console.log("Final File Path resolved:", filePath);
              
              if (!filePath) {
                if (rawFile) {
                  throw new Error("File terdeteksi tetapi informasi path tidak tersedia. Jika Anda menggunakan Electron versi terbaru, silakan coba gunakan tombol 'Pilih File Native' atau hubungi pengembang.");
                }
                throw new Error("Path file tidak ditemukan. Silakan coba unggah ulang file Anda.");
              }
              
              setProgress(30);
              // Get stats from Background Process instead of loading data here
              const stats = await ipcRenderer.invoke('get-join-stats', filePath);
              
              if (!stats.success) throw new Error(stats.error);

              console.log(`Joined stats (Background): ${stats.sheetCount} sheets, approx ${stats.totalRows} rows.`);
              
              setDataState(prev => ({ 
                ...prev, 
                data: [], // Keep empty to save memory in renderer
                originalData: [], 
                sheetCount: stats.sheetCount,
                rowCount: stats.totalRows
              }));
              setIsPreviewMode(true);
            } catch (err) {
              console.error("Join processing error:", err);
              setErrorInfo({
                title: "Gagal Memproses File",
                message: `${err.message}`
              });
            } finally {
              setIsGenerating(false);
            }
          }, 100);
          return;
        } else {
          // Download Joined Excel with Styling
          // Download Joined Excel with Styling
          setTimeout(async () => {
            try {
              if (activeTab === 'join-sheet') {
                // --- ULTIMATE BACKGROUND PATH: ELECTRON IPC ---
                // Renderer NEVER holds the data. Main process does everything.
                console.log("Processing Join via Background IPC...");
                
                const { ipcRenderer } = window.require('electron');
                const finalFileName = `Joined_Report_${new Date().toLocaleDateString().replace(/\//g, '-')}.xlsx`;

                // 1. Get save path from user
                const savePathResult = await ipcRenderer.invoke('show-save-dialog', {
                  title: 'Simpan Report Gabungan',
                  defaultPath: finalFileName,
                  filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
                });
                
                if (savePathResult.canceled || !savePathResult.filePath) {
                  setIsGenerating(false);
                  return;
                }
                
                setProgress(20);
                
                // 2. Process in Main process (Streaming)
                const { webUtils } = window.require('electron');
                let inputPath = dataState.rawFile?.path;
                
                if (!inputPath && webUtils && dataState.rawFile) {
                  inputPath = webUtils.getPathForFile(dataState.rawFile);
                }
                
                const outputPath = savePathResult.filePath;
                
                if (!inputPath) {
                  throw new Error("Path file asal tidak ditemukan. Silakan unggah ulang file.");
                }
                
                const result = await ipcRenderer.invoke('process-large-join', { inputPath, outputPath });
                
                if (result.success) {
                  setProgress(100);
                  alert("Report berhasil disimpan di: " + outputPath);
                } else {
                  throw new Error(result.error);
                }
              } else {
                // --- STYLIZED PATH: EXCELJS (For other tabs) ---
                const totalRows = dataState.data.length;
                const finalFileName = `Report_${new Date().toLocaleDateString().replace(/\//g, '-')}.xlsx`;
                const workbook = new ExcelJS.Workbook();
                const ws = workbook.addWorksheet("Report Data");
                
                const CHUNK_SIZE = 20000;
                for (let i = 0; i < totalRows; i += CHUNK_SIZE) {
                  const chunk = dataState.data.slice(i, i + CHUNK_SIZE);
                  ws.addRows(chunk);
                  setProgress(Math.min(Math.round(((i + chunk.length) / totalRows) * 50), 50));
                  await new Promise(res => setTimeout(res, 0));
                }

                const headerRow = ws.getRow(1);
                const borderStyle = {
                  top: { style: 'thin', color: { argb: 'FF000000' } },
                  left: { style: 'thin', color: { argb: 'FF000000' } },
                  bottom: { style: 'thin', color: { argb: 'FF000000' } },
                  right: { style: 'thin', color: { argb: 'FF000000' } }
                };

                const colCount = ws.columnCount;
                for (let c = 1; c <= colCount; c++) {
                  const column = ws.getColumn(c);
                  column.border = borderStyle;
                  if (c === 2) column.alignment = { vertical: 'middle', horizontal: 'center' };
                  setProgress(50 + Math.round((c / colCount) * 40));
                  await new Promise(res => setTimeout(res, 0));
                }

                headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
                headerRow.border = borderStyle;
                headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
                ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];

                for (let c = 1; c <= colCount; c++) {
                  const headerVal = headerRow.getCell(c).value;
                  const headerStr = headerVal ? headerVal.toString() : "";
                  ws.getColumn(c).width = Math.min(Math.max(headerStr.length + 5, 12), 40);
                  if (c % 10 === 0) {
                    setProgress(90 + Math.round((c / colCount) * 5));
                    await new Promise(res => setTimeout(res, 0));
                  }
                }

                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                saveAs(blob, finalFileName);
                setProgress(100);
              }
            } catch (err) {
              console.error(err);
              setErrorInfo({
                title: "Gagal Mengunduh File",
                message: `Kesalahan saat membuat file Excel: ${err.message}.`
              });
            } finally {
              setIsGenerating(false);
            }
          }, 100);
          return;
        }
      } else if (!isPreviewMode) {
        // Step 1: Generate Preview
        setTimeout(() => {
          try {
            const previewData = generatePreviewData(dataState.originalData);
            setDataState(prev => ({ ...prev, data: previewData }));
            setIsPreviewMode(true);
          } catch (err) {
            console.error(err);
            setErrorInfo({
              title: "Gagal Memproses Data",
              message: `Terjadi kesalahan saat memproses preview: ${err.message}.`
            });
          } finally {
            setIsGenerating(false);
          }
        }, 100);
      } else {
        // Step 2: Download Real Report
        setTimeout(async () => {
          try {
            const reportFileName = activeTab === 'report-sasaran' ? 'Report Performance Montly.xlsx' : dataState.fileName;
            await generateAHTReport(dataState.originalData, reportFileName);
          } catch (err) {
            console.error(err);
            setErrorInfo({
              title: "Gagal Mengunduh Report",
              message: `Terjadi kesalahan saat membuat report Excel: ${err.message}.`
            });
          } finally {
            setIsGenerating(false);
          }
        }, 100);
      }
    } catch (error) {
      console.error("Failed to generate report:", error);
      setErrorInfo({
        title: "Kesalahan Sistem",
        message: `Terjadi kesalahan tak terduga: ${error.message}.`
      });
      setIsGenerating(false);
    }
  };

  const handleInputToSummary = async () => {
    try {
      setIsGenerating(true);
      const blob = await processSummaryReport(dataState.originalData, dataState.files || (dataState.rawFile ? [dataState.rawFile] : []));
      saveAs(blob, "Summary Sasaran Mutu All Agent Montly_Updated.xlsx");
      
      // Trigger the standard report download as well so both are saved
      if (dataState.originalData) {
          const reportFileName = activeTab === 'report-sasaran' ? 'Report Performance Montly.xlsx' : dataState.fileName;
          await generateAHTReport(dataState.originalData, reportFileName);
      }
    } catch (error) {
      console.error(error);
      setErrorInfo({
        title: "Gagal Input ke Summary",
        message: `Kesalahan memproses summary: ${error.message}`
      });
    } finally {
      setIsGenerating(false);
    }
  };

  const handleExport = () => {
    const dataToExport = isPreviewMode ? displayData : dataState.data;
    if (!dataToExport) return;

    const ws = XLSX.utils.aoa_to_sheet(dataToExport, { dense: true });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    
    // Ensure .xlsx extension
    const baseName = dataState.fileName ? dataState.fileName.split('.')[0] : 'report';
    XLSX.writeFile(wb, `processed_${baseName}.xlsx`);
  };

  // Compute filtering for Preview Mode
  const { displayData, uniqueSheets } = React.useMemo(() => {
    if (!dataState.data || !isPreviewMode) return { displayData: dataState.data, uniqueSheets: [] };

    // If it's Join Sheet, don't apply the AHT-specific filtering logic
    if (activeTab === 'join-sheet') {
      return {
        displayData: dataState.data,
        uniqueSheets: []
      };
    }

    // dataState.data[0] is Header. rows are 1..n
    const headers = dataState.data[0];
    const rows = dataState.data.slice(1);

    // Skill is index 3 in preview data (AHT Report specific)
    const skills = [...new Set(rows.map(r => r[3]))].filter(Boolean).sort();

    let filteredRows = rows;
    if (previewSheet !== 'All') {
      filteredRows = rows.filter(r => r[3] === previewSheet);
    }

    return {
      displayData: [headers, ...filteredRows],
      uniqueSheets: skills
    };
  }, [dataState.data, isPreviewMode, previewSheet, activeTab]);

  // Get current active report label
  const activeReport = reportMenus.find(m => m.id === activeTab) || INITIAL_REPORT_MENUS[0];
  const isDashboard = activeTab === 'dashboard';

  return (
    <div className="h-screen flex flex-col text-slate-200 selection:bg-[#818cf8] selection:text-white bg-slate-900 overflow-hidden">
      {/* Background Ambience */}
      <div className="fixed inset-0 z-[-1] overflow-hidden pointer-events-none">
        <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-blue-600/20 rounded-full blur-[120px]" />
        <div className="absolute bottom-[-10%] right-[-10%] w-[40%] h-[40%] bg-purple-600/20 rounded-full blur-[120px]" />
      </div>

      {/* Navbar */}
      <nav className="shrink-0 border-b border-white/5 bg-slate-900/50 backdrop-blur-md z-30">
        <div className="container px-4 mx-auto flex items-center justify-between py-4">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-gradient-to-tr from-[#818cf8] to-purple-500 rounded-lg flex items-center justify-center shadow-lg shadow-[#818cf8]/20">
              <Database className="w-5 h-5 text-white" />
            </div>
            <h1 className="text-xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-white to-slate-400">
              Otomasi Report
            </h1>
          </div>

          <div className="flex items-center gap-4">
            <button className="p-2 hover:bg-white/5 rounded-lg transition-colors text-slate-400 hover:text-white">
              <Settings className="w-5 h-5" />
            </button>
            <button className="p-2 hover:bg-white/5 rounded-lg transition-colors text-slate-400 hover:text-white">
              <Github className="w-5 h-5" />
            </button>
          </div>
        </div>
      </nav>

      {/* Main Layout */}
      <div className="flex flex-1 overflow-hidden">
        {/* Sidebar */}
        <aside className="w-64 border-r border-white/5 bg-slate-900/30 backdrop-blur-sm shrink-0 flex flex-col overflow-y-auto custom-scrollbar">
          <div className="p-4 space-y-2">
            <div className="px-3 py-2 text-xs font-semibold text-slate-500 uppercase tracking-wider">
              Main Menu
            </div>

            <button
              onClick={() => handleNavClick('dashboard')}
              className={`w-full flex items-center gap-3 px-3 py-2 text-sm font-medium rounded-lg transition-colors ${activeTab === 'dashboard'
                ? 'bg-[#818cf8]/10 text-[#818cf8]'
                : 'text-slate-300 hover:bg-slate-800/50 hover:text-white'
                }`}
            >
              <LayoutDashboard className="w-4 h-4" />
              Dashboard
            </button>

            <div className="pt-4 px-3 py-2 text-xs font-semibold text-slate-500 uppercase tracking-wider flex items-center justify-between group">
              <span>Reports</span>
              <button 
                onClick={() => setShowAddModal(true)}
                className="p-1 hover:bg-white/10 rounded-md transition-colors text-[#818cf8]"
                title="Add New Report"
              >
                <Plus className="w-3.5 h-3.5" />
              </button>
            </div>

            {reportMenus.map((menu) => {
              const Icon = menu.icon;
              const isActive = activeTab === menu.id;
              return (
                <button
                  key={menu.id}
                  onClick={() => handleNavClick(menu.id)}
                  className={`w-full flex items-center gap-3 px-3 py-2 text-sm font-medium rounded-lg transition-colors text-left ${isActive
                    ? 'bg-[#818cf8]/10 text-[#818cf8]'
                    : 'text-slate-300 hover:bg-slate-800/50 hover:text-white'
                    }`}
                >
                  <Icon className="w-4 h-4 shrink-0" />
                  <span className="truncate">{menu.label}</span>
                </button>
              );
            })}
          </div>
        </aside>

        {/* Content Area */}
        <main className="flex-1 p-8 overflow-hidden flex flex-col relative">
          {isDashboard ? (
            <div className="text-center mt-20 flex-1 overflow-y-auto">
              <h2 className="text-3xl font-bold text-slate-200">Dashboard</h2>
              <p className="text-slate-400 mt-2">Welcome to your dashboard overview.</p>
            </div>
          ) : (
            <AnimatePresence mode='wait'>
              {!dataState.data ? (
                <motion.div
                  key="upload"
                  initial={{ opacity: 0, y: 20 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, scale: 0.95 }}
                  className="w-full h-full flex flex-col items-center justify-center gap-8"
                >
                  <div className="text-center space-y-4">
                    <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-[#818cf8]/10 text-[#818cf8] text-sm font-medium mb-2">
                      <Wand2 className="w-3 h-3" />
                      <span>{activeReport ? activeReport.label : 'Report'} Active</span>
                    </div>
                    <h2 className="text-4xl md:text-5xl font-bold tracking-tight">
                      Transform your Data <br />
                      <span className="text-transparent bg-clip-text bg-gradient-to-r from-[#818cf8] to-purple-400">
                        Instantly
                      </span>
                    </h2>
                    <p className="text-lg text-slate-400 max-w-lg mx-auto">
                      Upload your Excel files to generate the <strong>{activeReport ? activeReport.label : 'Report'}</strong> with one click.
                    </p>

                    {activeTab.startsWith('custom-') && (
                      <button 
                        onClick={() => alert("Customization feature coming soon!")}
                        className="mt-4 px-6 py-2 rounded-full border border-[#818cf8]/30 bg-[#818cf8]/5 text-[#818cf8] text-sm font-semibold hover:bg-[#818cf8]/10 transition-all flex items-center gap-2 mx-auto"
                      >
                        <Settings className="w-4 h-4" />
                        Customize Report Logic
                      </button>
                    )}
                  </div>

                  <div className={`w-full flex flex-col items-center gap-8 transition-all duration-500 ${activeTab.startsWith('custom-') ? 'blur-lg opacity-30' : ''}`}>
                    <FileUpload 
                      onDataLoaded={handleDataLoaded} 
                      activeTab={activeTab} 
                      onLoading={(val) => setIsGenerating(val)} 
                      disabled={activeTab.startsWith('custom-')}
                    />

                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 w-full mt-8">
                      {[
                        { title: "Fast Processing", desc: "Instant client-side parsing" },
                        { title: "Automation", desc: "Build workflows to clean data" },
                        { title: "Export Ready", desc: "Convert back to Excel easily" },
                      ].map((feature, i) => (
                        <div key={i} className="p-4 rounded-xl bg-slate-800/30 border border-slate-700/50 backdrop-blur-sm text-center">
                          <h3 className="font-semibold text-slate-200">{feature.title}</h3>
                          <p className="text-sm text-slate-500">{feature.desc}</p>
                        </div>
                      ))}
                    </div>
                  </div>

                  {activeTab.startsWith('custom-') && (
                    <button 
                      onClick={() => {
                        setMenuToDelete(activeReport);
                        setShowDeleteModal(true);
                      }}
                      className="mt-12 px-6 py-2 rounded-full border border-[#818cf8]/30 bg-[#818cf8]/5 text-[#818cf8] text-sm font-semibold hover:bg-[#818cf8]/10 transition-all flex items-center gap-2 mx-auto"
                    >
                      <Trash2 className="w-4 h-4" />
                      Hapus Menu Report ({activeReport.label})
                    </button>
                  )}
                </motion.div>
              ) : (
                <motion.div
                  key="table"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  className="w-full h-full flex flex-col min-h-0"
                >
                  {/* Sticky Header with Title and Controls */}
                  <div className="shrink-0 z-20 bg-slate-900/95 backdrop-blur-xl border-b border-white/10 pb-4 mb-4 -mx-8 px-8 pt-6 sticky top-0 shadow-2xl">
                    {/* Top Row: Report Title & File Info (Centered) */}
                    <div className="flex flex-col items-center justify-center mb-6 text-center space-y-1">
                      <div className="flex items-center gap-2">
                        <h2 className="text-xl font-bold text-white drop-shadow-md tracking-wide">
                          {activeReport ? activeReport.label : 'Data Preview'}
                        </h2>
                      </div>
                      <div className="flex items-center gap-2 text-sm">
                        <span className="font-semibold text-slate-300">{dataState.fileName}</span>
                        <span className="text-slate-500">•</span>
                        <div className="flex items-center gap-1 text-slate-400">
                          <Filter className="w-3 h-3" />
                          <span>
                            {isPreviewMode && displayData ? displayData.length - 1 : (dataState.data ? dataState.data.length - 1 : 0)} rows found
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* Toolbar Row */}
                    <div className="flex flex-wrap items-center justify-between gap-4 w-full px-2">
                      {/* Left Group: Search & Sheet */}
                      <div className="flex items-center gap-3 flex-1 min-w-[200px]">
                        {/* Search Pill */}
                        <div className="relative group rounded-full shadow-lg shadow-cyan-500/20 flex-1 max-w-md">
                          <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none text-slate-900/50">
                            <Search className="w-4 h-4" />
                          </div>
                          <input
                            type="text"
                            placeholder="Search Data..."
                            value={searchTerm}
                            onChange={(e) => setSearchTerm(e.target.value)}
                            className="block w-full pl-10 pr-6 h-12 border-none rounded-full bg-gradient-to-r from-cyan-400 to-blue-500 text-slate-900 font-bold placeholder-slate-800/50 focus:outline-none focus:ring-2 focus:ring-cyan-300 transition-all text-sm"
                          />
                        </div>

                        {/* Sheet Selection Pill */}
                        {isPreviewMode && (
                          <div className="relative rounded-full shadow-lg shadow-cyan-500/20 shrink-0">
                            <div className="absolute inset-y-0 right-0 pr-3 flex items-center pointer-events-none text-slate-900">
                              <svg className="h-4 w-4 fill-current" viewBox="0 0 20 20"><path d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z" /></svg>
                            </div>
                            <select
                              value={previewSheet}
                              onChange={(e) => setPreviewSheet(e.target.value)}
                              className="appearance-none pl-5 pr-10 h-12 rounded-full bg-gradient-to-r from-cyan-400 to-blue-500 text-slate-900 font-bold text-sm focus:outline-none cursor-pointer transition-colors border-none ring-0"
                            >
                              <option value="All">All Sheets</option>
                              {uniqueSheets.map(sheet => (
                                <option key={sheet} value={sheet}>{sheet}</option>
                              ))}
                            </select>
                          </div>
                        )}
                      </div>

                      {/* Right Group: Actions */}
                      <div className="flex items-center gap-3 shrink-0">
                        <button
                          onClick={handleReset}
                          className="h-12 px-6 rounded-full shadow-lg shadow-cyan-500/20 bg-gradient-to-r from-cyan-400 to-blue-500 text-slate-900 text-sm font-bold transition-all hover:brightness-110 flex items-center justify-center hover:scale-105"
                        >
                          Close
                        </button>

                        {isPreviewMode && activeTab === 'report-sasaran' && (
                          <button
                            onClick={handleInputToSummary}
                            disabled={isGenerating}
                            className="h-12 px-6 rounded-full shadow-lg shadow-green-500/20 bg-gradient-to-r from-green-400 to-emerald-500 text-slate-900 text-sm font-bold transition-all hover:brightness-110 flex items-center gap-2 hover:scale-105"
                          >
                            <FileSpreadsheet className={`w-4 h-4 ${isGenerating ? 'animate-spin' : ''}`} />
                            <span>Input ke Summary</span>
                          </button>
                        )}

                        <button
                          onClick={handleGenerateReport}
                          disabled={isGenerating}
                          className="h-12 px-6 rounded-full shadow-lg shadow-cyan-500/20 bg-gradient-to-r from-cyan-400 to-blue-500 text-slate-900 text-sm font-bold transition-all hover:brightness-110 flex items-center gap-2 hover:scale-105"
                        >
                          <Wand2 className={`w-4 h-4 ${isGenerating ? 'animate-spin' : ''}`} />
                          <span>
                            {isGenerating
                              ? 'Processing...'
                              : (isPreviewMode ? 'Download Report' : (activeTab === 'join-sheet' ? 'Gabungkan Sheet' : 'Otomasi Report'))
                            }
                          </span>
                        </button>
                      </div>
                    </div>
                  </div>

                   <div className="flex-1 overflow-y-auto custom-scrollbar rounded-xl border border-white/5 bg-slate-900/30 backdrop-blur-sm p-1">
                    {isPreviewMode && activeTab === 'join-sheet' ? (
                      <div className="h-full flex flex-col items-center justify-center p-8 text-center gap-4">
                        <div className="w-16 h-16 rounded-full bg-cyan-500/10 flex items-center justify-center ring-1 ring-cyan-500/20">
                          <FileSpreadsheet className="w-8 h-8 text-cyan-400" />
                        </div>
                        <div className="space-y-2">
                          <h3 className="text-xl font-bold text-white">Data Siap Diunduh</h3>
                          <p className="text-slate-400 max-w-md mx-auto text-sm">
                            Berhasil menganalisa <strong>{dataState.sheetCount}</strong> sheet dengan total sekitar <strong>{dataState.rowCount}</strong> baris. 
                            Preview dinonaktifkan untuk menghemat memori demi stabilitas aplikasi. 
                            Silakan klik <strong>Download Report</strong> untuk menyimpan hasil penggabungan.
                          </p>
                        </div>
                      </div>
                    ) : (
                      <DataTable
                        data={isPreviewMode ? displayData : dataState.data}
                        fileName={dataState.fileName}
                        searchTerm={searchTerm}
                      />
                    )}
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          )}
        </main>
      </div>

      <div className="fixed bottom-0 left-0 w-64 p-4 hidden md:block">
        <div className="text-xs text-slate-600 text-center">
          &copy; {new Date().getFullYear()} Excel Processor
        </div>
      </div>

      {/* Navigation Confirmation Modal */}
      <AnimatePresence>
        {showNavConfirm && (
          <div className="fixed top-0 left-0 w-screen h-screen z-[9999] flex items-center justify-center bg-slate-950/80 backdrop-blur-xl">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              style={{ backgroundColor: '#0f172a' }} // Force dark background
              className="w-full max-w-[340px] border border-white/10 ring-1 ring-white/5 rounded-2xl shadow-2xl shadow-black/80 overflow-hidden relative"
            >
              <div className="absolute inset-0 bg-gradient-to-br from-white/5 to-transparent pointer-events-none" />
              <div className="p-6 flex flex-col items-center text-center relative z-10">
                <div className="w-14 h-14 rounded-full bg-amber-500/10 flex items-center justify-center mb-4 ring-1 ring-amber-500/20 shadow-[0_0_15px_rgba(245,158,11,0.2)]">
                  <AlertTriangle className="w-7 h-7 text-amber-500" />
                </div>
                <h3 className="text-xl font-bold text-white mb-2 tracking-tight">Switch Menu?</h3>
                <p className="text-sm text-slate-400 mb-6 leading-relaxed px-2">
                  You have unsaved changes. Switching menus will <span className="text-amber-500 font-medium">clear your current data</span>.
                </p>
                <div className="flex gap-3 w-full">
                  <button
                    onClick={cancelNavigation}
                    className="flex-1 px-4 py-2.5 rounded-xl border border-white/10 bg-white/5 text-slate-300 text-sm font-semibold hover:bg-white/10 hover:text-white transition-all"
                  >
                    Cancel
                  </button>
                  <button
                    onClick={confirmNavigation}
                    className="flex-1 px-4 py-2.5 rounded-xl bg-gradient-to-r from-amber-500 to-orange-600 text-white text-sm font-bold hover:brightness-110 hover:shadow-lg hover:shadow-amber-500/20 transition-all transform hover:scale-[1.02]"
                  >
                    Switch
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
      {/* Month Selection Modal */}
      <AnimatePresence>
        {showMonthSelector && (
          <div className="fixed top-0 left-0 w-screen h-screen z-[9999] flex items-center justify-center bg-slate-950/80 backdrop-blur-xl">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              style={{ backgroundColor: '#0f172a' }}
              className="w-full max-w-[400px] border border-white/10 ring-1 ring-white/5 rounded-2xl shadow-2xl shadow-black/80 overflow-hidden relative"
            >
              <div className="absolute inset-0 bg-gradient-to-br from-white/5 to-transparent pointer-events-none" />
              <div className="p-6 flex flex-col items-center relative z-10">
                <h3 className="text-xl font-bold text-white mb-2 tracking-tight">Select Target Month</h3>
                <p className="text-sm text-slate-400 mb-6 text-center">
                  Sheet bulan apa yang ingin diisi?
                </p>

                <div className="grid grid-cols-3 gap-3 w-full mb-6">
                  {MONTH_OPTIONS.map(m => (
                    <button
                      key={m}
                      onClick={() => setSelectedMonth(m)}
                      className={`px-3 py-2 rounded-lg text-sm font-semibold transition-all ${selectedMonth === m
                        ? 'bg-[#818cf8] text-white shadow-lg shadow-[#818cf8]/25 scale-105'
                        : 'bg-slate-800 text-slate-400 hover:bg-slate-700 hover:text-slate-200'
                        }`}
                    >
                      {m}
                    </button>
                  ))}
                </div>

                <div className="flex gap-3 w-full">
                  <button
                    onClick={() => setShowMonthSelector(false)}
                    className="flex-1 px-4 py-2.5 rounded-xl border border-white/10 bg-white/5 text-slate-300 text-sm font-semibold hover:bg-white/10 hover:text-white transition-all"
                  >
                    Cancel
                  </button>
                  <button
                    onClick={handleMonthConfirm}
                    className="flex-1 px-4 py-2.5 rounded-xl bg-gradient-to-r from-cyan-500 to-blue-600 text-white text-sm font-bold hover:brightness-110 hover:shadow-lg hover:shadow-cyan-500/20 transition-all transform hover:scale-[1.02]"
                  >
                    Process
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
      {/* Global Loading Pop-up */}
      <AnimatePresence>
        {isGenerating && (
          <div className="fixed inset-0 w-full h-full z-[10000] flex flex-col items-center justify-center bg-slate-950/50 backdrop-blur-sm overflow-hidden p-4">
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              style={{ backgroundColor: '#0f172a' }}
              className="w-full max-w-[440px] border border-white/10 ring-1 ring-white/5 rounded-3xl shadow-[0_32px_64px_-12px_rgba(0,0,0,0.8)] overflow-hidden relative"
            >
              {/* Progress Line */}
              <div className="absolute top-0 left-0 w-full h-1.5 bg-slate-800">
                <motion.div 
                  initial={{ width: "0%" }}
                  animate={{ width: `${progress}%` }}
                  transition={{ duration: 0.5 }}
                  className="h-full bg-gradient-to-r from-[#818cf8] to-purple-500"
                />
              </div>

              <div className="p-12 flex flex-col items-center text-center relative z-10 gap-8">
                {/* 1. Processing Data (Title) */}
                <motion.h3
                  initial={{ opacity: 0, y: -10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="text-2xl font-bold text-white m-0 tracking-tight"
                >
                  {isPreviewMode ? "Generating Report" : "Processing Data"}
                </motion.h3>

                {/* 2. Original Custom SVG Icon Loading (Smaller & Always Spinning) */}
                <div className="relative m-0 flex items-center justify-center">
                  <div className="w-12 h-12 flex items-center justify-center animate-spin-infinite">
                    <svg width="48" height="48" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
                      <path d="M50 27.5c-2.2 0-4-1.8-4-4v-12c0-2.2 1.8-4 4-4s4 1.8 4 4v12c0 2.2-1.8 4-4 4z" fill="#d65a62"/>
                      <path d="M38.8 30.5c-1.9 1.1-4.4.4-5.5-1.5l-6-10.4c-1.1-1.9-.4-4.4 1.5-5.5 1.9-1.1 4.4-.4 5.5 1.5l6 10.4c1 2 .4 4.4-1.5 5.5z" fill="#ea7c60"/>
                      <path d="M30.5 38.7c-1.1 1.9-3.6 2.6-5.5 1.5l-10.4-6c-1.9-1.1-2.6-3.6-1.5-5.5s3.6-2.6 5.5-1.5l10.4 6c2 1.2 2.6 3.6 1.5 5.5z" fill="#f0af6b"/>
                      <path d="M27.5 50c0 2.2-1.8 4-4 4h-12c-2.2 0-4-1.8-4-4s1.8-4 4-4h12c2.2 0 4 1.8 4 4z" fill="#d65a62"/>
                      <path d="M30.5 61.2c1.1 1.9.4 4.4-1.5 5.5l-10.4 6c-1.9 1.1-4.4.4-5.5-1.5-1.1-1.9-.4-4.4 1.5-5.5l10.4-6c2-1 4.4-.4 5.5 1.5z" fill="#ea7c60"/>
                      <path d="M38.7 69.5c1.9 1.1 2.6 3.6 1.5 5.5l-6 10.4c-1.1 1.9-3.6 2.6-5.5 1.5s-2.6-3.6-1.5-5.5l6-10.4c1.2-2 3.6-2.6 5.5-1.5z" fill="#f0af6b"/>
                      <path d="M50 72.5c2.2 0 4 1.8 4 4v12c0 2.2-1.8 4-4 4s-4-1.8-4-4v-12c0-2.2 1.8-4 4-4z" fill="#d65a62"/>
                      <path d="M61.2 69.5c1.9-1.1 4.4-.4 5.5 1.5l6 10.4c1.1 1.9.4 4.4-1.5 5.5-1.9 1.1-4.4.4-5.5-1.5l-6-10.4c-1-2-.4-4.4 1.5-5.5z" fill="#ea7c60"/>
                      <path d="M69.5 61.3c1.1-1.9 3.6-2.6 5.5-1.5l10.4 6c1.9 1.1 2.6 3.6 1.5 5.5s-3.6 2.6-5.5 1.5l-10.4-6c-2-1.2-2.6-3.6-1.5-5.5z" fill="#f0af6b"/>
                      <path d="M72.5 50c0-2.2 1.8-4 4-4h12c2.2 0 4 1.8 4 4s-1.8 4-4 4h-12c-2.2 0-4-1.8-4-4z" fill="#d65a62"/>
                      <path d="M69.5 38.8c-1.1-1.9-.4-4.4 1.5-5.5l10.4-6c1.9-1.1 4.4-.4 5.5 1.5 1.1 1.9.4 4.4-1.5 5.5l-10.4 6c-2 1-4.4.4-5.5-1.5z" fill="#ea7c60"/>
                      <path d="M61.3 30.5c-1.9-1.1-2.6-3.6-1.5-5.5l6-10.4c1.1-1.9 3.6-2.6 5.5-1.5s2.6 3.6 1.5 5.5l-6 10.4c-1.2 2-3.6 2.6-5.5 1.5z" fill="#f0af6b"/>
                    </svg>
                  </div>
                </div>

                {/* 3. Description & Footer */}
                <motion.div
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="flex flex-col items-center gap-4 m-0"
                >
                  <p className="text-slate-400 text-sm leading-relaxed px-8">
                    {isPreviewMode 
                      ? "Sedang menyiapkan file Excel hasil pengolahan data Anda." 
                      : "Mohon tunggu sebentar, sistem sedang memproses data dalam jumlah besar agar hasilnya akurat."
                    }
                  </p>
                  <div className="pt-2">
                    <p className="text-[#818cf8] text-[10px] font-bold uppercase tracking-[0.2em] animate-pulse">
                      Sistem sedang bekerja
                    </p>
                  </div>
                </motion.div>

                {/* Animated Dots */}
                <div className="flex gap-2 justify-center mt-8">
                  {[0, 1, 2].map((i) => (
                    <motion.div
                      key={i}
                      animate={{ 
                        opacity: [0.3, 1, 0.3],
                        scale: [1, 1.2, 1],
                        y: [0, -4, 0]
                      }}
                      transition={{ duration: 1.5, repeat: Infinity, delay: i * 0.2 }}
                      className="w-2.5 h-2.5 rounded-full bg-[#818cf8] shadow-[0_0_10px_rgba(129,140,248,0.5)]"
                    />
                  ))}
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Success Modal */}
      <AnimatePresence>
        {showSuccess && (
          <div className="fixed top-0 left-0 w-screen h-screen z-[10002] flex items-center justify-center bg-slate-950/80 backdrop-blur-xl">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              style={{ backgroundColor: '#0f172a' }}
              className="w-full max-w-[400px] border border-green-500/20 ring-1 ring-green-500/10 rounded-2xl shadow-2xl shadow-black/80 overflow-hidden relative"
            >
              <div className="absolute inset-0 bg-gradient-to-br from-green-500/5 to-transparent pointer-events-none" />
              <div className="p-8 flex flex-col items-center text-center relative z-10">
                <div className="w-16 h-16 rounded-full bg-green-500/10 flex items-center justify-center mb-6 ring-1 ring-green-500/20 shadow-[0_0_20px_rgba(34,197,94,0.2)]">
                  <motion.div
                    initial={{ scale: 0 }}
                    animate={{ scale: 1 }}
                    transition={{ type: "spring", damping: 12, stiffness: 200, delay: 0.1 }}
                  >
                    <svg className="w-8 h-8 text-green-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={3}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                    </svg>
                  </motion.div>
                </div>
                <h3 className="text-2xl font-bold text-white mb-2 tracking-tight">Berhasil!</h3>
                <p className="text-slate-400 mb-8 leading-relaxed">
                  {successMessage}
                </p>
                <button
                  onClick={() => setShowSuccess(false)}
                  className="w-full px-6 py-3 rounded-xl bg-gradient-to-r from-green-500 to-emerald-600 text-white font-bold hover:brightness-110 hover:shadow-lg hover:shadow-green-500/20 transition-all transform hover:scale-[1.02]"
                >
                  Tutup
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Delete Report Modal */}
      <AnimatePresence>
        {showDeleteModal && (
          <div className="fixed top-0 left-0 w-screen h-screen z-[9999] flex items-center justify-center bg-slate-950/80 backdrop-blur-xl">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              style={{ backgroundColor: '#0f172a' }}
              className="w-full max-w-[400px] border border-red-500/20 ring-1 ring-red-500/10 rounded-2xl shadow-2xl shadow-black/80 overflow-hidden relative"
            >
              <div className="absolute inset-0 bg-gradient-to-br from-red-500/5 to-transparent pointer-events-none" />
              <div className="p-8 flex flex-col relative z-10">
                <div className="w-14 h-14 rounded-full bg-red-500/10 flex items-center justify-center mb-6 ring-1 ring-red-500/20 shadow-[0_0_20px_rgba(239,68,68,0.2)] mx-auto">
                  <Trash2 className="w-7 h-7 text-red-500" />
                </div>
                
                <h3 className="text-2xl font-bold text-white mb-2 tracking-tight text-center">Delete Menu</h3>
                <p className="text-sm text-slate-400 mb-8 text-center leading-relaxed">
                  Enter password to delete <strong>{menuToDelete?.label}</strong>.
                </p>

                <div className="space-y-4 mb-8">
                  <input
                    ref={deleteInputRef}
                    type="password"
                    placeholder={lockoutUntil ? "Locked Out" : "Password"}
                    value={deletePassword}
                    disabled={!!lockoutUntil}
                    onChange={(e) => {
                      setDeletePassword(e.target.value);
                      if (deleteError) setDeleteError(null);
                    }}
                    onKeyDown={(e) => e.key === 'Enter' && handleDeleteMenu()}
                    className={`w-full px-4 py-3 rounded-xl bg-slate-800/50 border transition-all text-center text-white placeholder-slate-500 focus:outline-none focus:ring-2 ${deleteError ? 'border-red-500 focus:ring-red-500/50' : 'border-white/10 focus:ring-red-500/50'} ${lockoutUntil ? 'opacity-50 cursor-not-allowed' : ''}`}
                  />
                  {(deleteError || lockoutUntil) && (
                    <motion.div 
                      initial={{ opacity: 0, y: -5 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="space-y-1 text-center"
                    >
                      <p className="text-xs text-red-500 font-medium">
                        {lockoutUntil 
                          ? "Password Salah 3X, Tunggu 15 menit untuk mencoba kembali"
                          : deleteError
                        }
                      </p>
                      {lockoutUntil && (
                        <p className="text-[10px] text-slate-500 font-bold uppercase tracking-wider">
                          Sisa waktu: {Math.max(0, Math.floor((lockoutUntil - currentTime) / 60000))}m {Math.max(0, Math.floor(((lockoutUntil - currentTime) % 60000) / 1000))}s
                        </p>
                      )}
                    </motion.div>
                  )}
                </div>

                <div className="flex gap-3 w-full">
                  <button
                    onClick={() => {
                      setShowDeleteModal(false);
                      setDeletePassword('');
                    }}
                    className="flex-1 px-4 py-3 rounded-xl border border-white/10 bg-white/5 text-slate-300 text-sm font-semibold hover:bg-white/10 hover:text-white transition-all"
                  >
                    Cancel
                  </button>
                  <button
                    onClick={handleDeleteMenu}
                    disabled={!deletePassword || !!lockoutUntil}
                    className="flex-1 px-4 py-3 rounded-xl bg-gradient-to-r from-red-500 to-red-600 text-white text-sm font-bold hover:brightness-110 hover:shadow-lg hover:shadow-red-500/20 transition-all transform hover:scale-[1.02] disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none"
                  >
                    Confirm Delete
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Add New Report Modal */}
      <AnimatePresence>
        {showAddModal && (
          <div className="fixed top-0 left-0 w-screen h-screen z-[9999] flex items-center justify-center bg-slate-950/80 backdrop-blur-xl">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              style={{ backgroundColor: '#0f172a' }}
              className="w-full max-w-[400px] border border-white/10 ring-1 ring-white/5 rounded-2xl shadow-2xl shadow-black/80 overflow-hidden relative"
            >
              <div className="absolute inset-0 bg-gradient-to-br from-white/5 to-transparent pointer-events-none" />
              <div className="p-8 flex flex-col relative z-10">
                <div className="w-14 h-14 rounded-full bg-[#818cf8]/10 flex items-center justify-center mb-6 ring-1 ring-[#818cf8]/20 shadow-[0_0_20px_rgba(129,140,248,0.2)] mx-auto">
                  <Plus className="w-7 h-7 text-[#818cf8]" />
                </div>
                
                <h3 className="text-2xl font-bold text-white mb-2 tracking-tight text-center">Add New Report</h3>
                <p className="text-sm text-slate-400 mb-8 text-center leading-relaxed">
                  Enter a name for your new custom report menu.
                </p>

                <div className="space-y-4 mb-8">
                  <div className="relative group">
                    <input
                      type="text"
                      placeholder="Report Name (e.g. Monthly Sales)"
                      value={newMenuName}
                      onChange={(e) => setNewMenuName(e.target.value)}
                      autoFocus
                      className="w-full px-4 py-3 rounded-xl bg-slate-800/50 border border-white/10 text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-[#818cf8] transition-all"
                    />
                  </div>
                  <div className="relative group">
                    <input
                      type="password"
                      placeholder="Delete Password"
                      value={newMenuPassword}
                      onChange={(e) => setNewMenuPassword(e.target.value)}
                      onKeyDown={(e) => e.key === 'Enter' && handleAddReport()}
                      className="w-full px-4 py-3 rounded-xl bg-slate-800/50 border border-white/10 text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-[#818cf8] transition-all"
                    />
                    <p className="text-[10px] text-slate-500 mt-1 ml-1">Password ini digunakan saat ingin menghapus menu ini nanti.</p>
                  </div>
                </div>

                <div className="flex gap-3 w-full">
                  <button
                    onClick={() => {
                      setShowAddModal(false);
                      setNewMenuName('');
                      setNewMenuPassword('');
                    }}
                    className="flex-1 px-4 py-3 rounded-xl border border-white/10 bg-white/5 text-slate-300 text-sm font-semibold hover:bg-white/10 hover:text-white transition-all"
                  >
                    Cancel
                  </button>
                  <button
                    onClick={handleAddReport}
                    disabled={!newMenuName.trim() || !newMenuPassword.trim()}
                    className="flex-1 px-4 py-3 rounded-xl bg-gradient-to-r from-[#818cf8] to-purple-600 text-white text-sm font-bold hover:brightness-110 hover:shadow-lg hover:shadow-[#818cf8]/20 transition-all transform hover:scale-[1.02] disabled:from-slate-800 disabled:to-slate-800 disabled:text-slate-500 disabled:cursor-not-allowed disabled:transform-none disabled:border disabled:border-white/5"
                  >
                    Create Menu
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Error Info Modal */}
      <AnimatePresence>
        {errorInfo && (
          <div className="fixed top-0 left-0 w-screen h-screen z-[10001] flex items-center justify-center bg-slate-950/80 backdrop-blur-xl">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              style={{ backgroundColor: '#0f172a' }}
              className="w-full max-w-[450px] border border-red-500/20 ring-1 ring-red-500/10 rounded-2xl shadow-2xl shadow-black/80 overflow-hidden relative"
            >
              <div className="absolute inset-0 bg-gradient-to-br from-red-500/5 to-transparent pointer-events-none" />
              <div className="p-8 flex flex-col items-center text-center relative z-10">
                <div className="w-16 h-16 rounded-full bg-red-500/10 flex items-center justify-center mb-6 ring-1 ring-red-500/20 shadow-[0_0_20px_rgba(239,68,68,0.2)]">
                  <AlertTriangle className="w-8 h-8 text-red-500" />
                </div>
                <h3 className="text-2xl font-bold text-white mb-3 tracking-tight">{errorInfo.title}</h3>
                <p className="text-slate-400 mb-8 leading-relaxed">
                  {errorInfo.message}
                </p>
                <button
                  onClick={() => setErrorInfo(null)}
                  className="w-full px-6 py-3 rounded-xl bg-gradient-to-r from-red-500 to-red-600 text-white font-bold hover:brightness-110 hover:shadow-lg hover:shadow-red-500/20 transition-all transform hover:scale-[1.02]"
                >
                  Dimengerti
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div >
  );
}

export default App;
