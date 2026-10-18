import React, { useRef, useState } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { Upload, FileSpreadsheet, X } from 'lucide-react';
import { cn } from '../lib/utils';
import * as XLSX from 'xlsx';
import { SheetSelector } from './SheetSelector';

export function FileUpload({ onDataLoaded, activeTab, onLoading, disabled }) {
    const [dragActive, setDragActive] = useState(false);
    const [error, setError] = useState(null);
    const [sheetSelection, setSheetSelection] = useState({ isOpen: false, sheets: [], workbook: null, fileName: null, allFiles: null });
    const inputRef = useRef(null);

    const handleDrag = (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (disabled) return;
        if (e.type === "dragenter" || e.type === "dragover") {
            setDragActive(true);
        } else if (e.type === "dragleave") {
            setDragActive(false);
        }
    };

    const processWorkbookSheet = (workbook, sheetName, fileName, rawFile, allFiles = null) => {
        try {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            onDataLoaded({
                fileName: allFiles ? `${allFiles.length} files selected` : fileName,
                data: jsonData,
                rawFile: rawFile,
                files: allFiles,
                multiple: allFiles !== null
            });
            // Reset selection state
            setSheetSelection({ isOpen: false, sheets: [], workbook: null, fileName: null, allFiles: null });
        } catch (err) {
            console.error(err);
            setError("Failed to parse the selected sheet.");
        }
    };

    const handleFiles = (filesList) => {
        setError(null);
        if (!filesList || filesList.length === 0) return;

        const files = Array.from(filesList);

        // Find the main data file (ignore Summary file if possible)
        let mainFile = files[0];
        if (files.length > 1) {
            const nonSummaryFile = files.find(f => !f.name.toLowerCase().includes('summary'));
            if (nonSummaryFile) {
                mainFile = nonSummaryFile;
            }
        }

        if (!mainFile.name.match(/\.(xlsx|xls|csv)$/)) {
            setError("Please upload a valid Excel file (.xlsx, .xls) or CSV.");
            return;
        }

        if (onLoading) onLoading(true);
        // Small timeout to allow UI to render spinner
        setTimeout(() => {
            processSingleFile(mainFile, files.length > 1 ? files : null);
        }, 100);
    };

    const processSingleFile = async (file, allFiles = null) => {
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);

            if (activeTab !== 'join-sheet' && workbook.SheetNames.length > 1) {
                setSheetSelection({
                    isOpen: true,
                    sheets: workbook.SheetNames,
                    workbook: workbook,
                    fileName: file.name,
                    rawFile: file, // Keep reference
                    allFiles: allFiles
                });
                if (onLoading) onLoading(false);
            } else {
                processWorkbookSheet(workbook, workbook.SheetNames[0], file.name, file, allFiles);
                if (onLoading) onLoading(false);
            }
        } catch (err) {
            console.error(err);
            setError("Failed to parse the file. Please try again.");
            if (onLoading) onLoading(false);
        }
    };

    const handleSheetSelect = (sheetName) => {
        if (sheetSelection.workbook) {
            processWorkbookSheet(sheetSelection.workbook, sheetName, sheetSelection.fileName, sheetSelection.rawFile, sheetSelection.allFiles);
        }
    };

    const handleDrop = (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (disabled) return;
        setDragActive(false);
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleFiles(e.dataTransfer.files);
        }
    };

    const handleChange = (e) => {
        e.preventDefault();
        if (e.target.files && e.target.files.length > 0) {
            handleFiles(e.target.files);
        }
    };

    const onButtonClick = () => {
        if (disabled) return;
        inputRef.current?.click();
    };

    const handleNativeSelect = async () => {
        try {
            if (typeof window.require !== 'function') return;
            const { ipcRenderer } = window.require('electron');
            const result = await ipcRenderer.invoke('open-file-dialog', {
                properties: ['openFile', 'multiSelections'],
                filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls', 'csv'] }]
            });

            if (!result.canceled && result.filePaths.length > 0) {
                const fs = window.require('fs');
                const fileObjs = result.filePaths.map(path => {
                    const buffer = fs.readFileSync(path);
                    const fileName = path.split(/[\\/]/).pop();
                    return {
                        path: path,
                        name: fileName,
                        size: buffer.length,
                        arrayBuffer: async () => buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength)
                    };
                });

                handleFiles(fileObjs);
            }
        } catch (err) {
            console.error("Native select error:", err);
            setError("Gagal menggunakan pemilih file native.");
        }
    };

    return (
        <div className="w-full max-w-xl mx-auto mb-8 space-y-4">
            <motion.div
                layout
                className={cn(
                    "relative border-2 border-dashed rounded-xl p-10 text-center transition-colors duration-200 ease-in-out glass-card",
                    disabled ? "cursor-not-allowed opacity-50 border-slate-700" : "cursor-pointer border-slate-600 hover:border-slate-500",
                    dragActive && !disabled ? "border-[#818cf8] bg-[#1e293b]/50" : "",
                    error ? "border-red-500/50" : ""
                )}
                onDragEnter={handleDrag}
                onDragLeave={handleDrag}
                onDragOver={handleDrag}
                onDrop={handleDrop}
                onClick={onButtonClick}
                whileHover={disabled ? {} : { scale: 1.01 }}
                whileTap={disabled ? {} : { scale: 0.99 }}
            >
                <input
                    ref={inputRef}
                    className="hidden"
                    type="file"
                    multiple
                    accept=".xlsx,.xls,.csv"
                    onChange={handleChange}
                />

                <div className="flex flex-col items-center justify-center gap-4">
                    <div className={cn(
                        "p-4 rounded-full bg-slate-800/50 ring-1 ring-white/10 shadow-xl",
                        dragActive && "animate-pulse"
                    )}>
                        <div className="relative">
                            <Upload className="w-8 h-8 text-[#818cf8]" />
                            {dragActive && (
                                <motion.div
                                    initial={{ opacity: 0, scale: 0.5 }}
                                    animate={{ opacity: 1, scale: 1 }}
                                    className="absolute -top-1 -right-1 w-3 h-3 bg-green-500 rounded-full"
                                />
                            )}
                        </div>
                    </div>

                    <div className="space-y-1">
                        <h3 className="text-lg font-semibold text-slate-200">
                            Drop your Excel file here
                        </h3>
                        <p className="text-sm text-slate-400">
                            or click to browse from your computer
                        </p>
                    </div>

                    <div className="flex items-center gap-2 text-xs text-slate-500 mt-2 bg-slate-900/40 px-3 py-1.5 rounded-full">
                        <FileSpreadsheet className="w-3.5 h-3.5" />
                        <span>Supports .xlsx, .xls, .csv</span>
                    </div>
                </div>
            </motion.div>

            {/* Native Fallback Button */}
            {typeof window.require === 'function' && (
                <div className="flex justify-center">
                    <button
                        onClick={(e) => {
                            e.stopPropagation();
                            if (disabled) return;
                            handleNativeSelect();
                        }}
                        disabled={disabled}
                        className={cn(
                            "text-xs underline transition-colors",
                            disabled ? "text-slate-600 no-underline cursor-not-allowed" : "text-[#818cf8] hover:text-[#818cf8]/80"
                        )}
                    >
                        Masalah dengan upload? Gunakan Pemilih File Native
                    </button>
                </div>
            )}

            <AnimatePresence>
                {error && (
                    <motion.div
                        initial={{ opacity: 0, y: -10 }}
                        animate={{ opacity: 1, y: 0 }}
                        exit={{ opacity: 0, y: -10 }}
                        className="mt-4 p-4 rounded-lg bg-red-500/10 border border-red-500/20 text-red-400 flex items-center justify-between"
                    >
                        <span className="text-sm">{error}</span>
                        <button
                            onClick={(e) => { e.stopPropagation(); setError(null); }}
                            className="p-1 hover:bg-red-500/20 rounded-full transition-colors"
                        >
                            <X className="w-4 h-4" />
                        </button>
                    </motion.div>
                )}
            </AnimatePresence>

            <SheetSelector
                isOpen={sheetSelection.isOpen}
                onClose={() => setSheetSelection({ ...sheetSelection, isOpen: false })}
                sheets={sheetSelection.sheets}
                onSelect={handleSheetSelect}
            />
        </div>
    );
}
