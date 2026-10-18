import React from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { FileSpreadsheet, X, Check } from 'lucide-react';

export function SheetSelector({ isOpen, onClose, sheets, onSelect }) {
    if (!isOpen) return null;

    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 0.95 }}
                className="w-full max-w-md bg-slate-900 border border-slate-700/50 rounded-2xl shadow-2xl overflow-hidden"
            >
                <div className="p-6 border-b border-slate-700/50 flex justify-between items-center bg-slate-800/50">
                    <div className="flex items-center gap-3">
                        <div className="p-2 bg-emerald-500/10 rounded-lg">
                            <FileSpreadsheet className="w-5 h-5 text-emerald-500" />
                        </div>
                        <h3 className="text-lg font-semibold text-slate-200">Select Sheet</h3>
                    </div>
                    <button
                        onClick={onClose}
                        className="p-2 hover:bg-white/5 rounded-lg text-slate-400 hover:text-white transition-colors"
                    >
                        <X className="w-5 h-5" />
                    </button>
                </div>

                <div className="flex-1 overflow-y-auto custom-scrollbar p-4" style={{ maxHeight: '60vh' }}>
                    <p className="text-sm text-slate-400 mb-4 px-2">
                        Found {sheets.length} sheets. Please select the one you want to process.
                    </p>
                    <div className="space-y-2 pb-4">
                        {sheets.map((sheet, idx) => (
                            <button
                                key={idx}
                                onClick={() => onSelect(sheet)}
                                className="w-full flex items-center justify-between p-4 rounded-xl bg-slate-800/30 border border-slate-700/30 hover:bg-slate-800 hover:border-slate-600 transition-all group text-left"
                            >
                                <div className="flex items-center gap-3">
                                    <span className="w-8 h-8 rounded-full bg-slate-700 flex items-center justify-center text-xs font-medium text-slate-400 group-hover:bg-[#818cf8] group-hover:text-white transition-colors">
                                        {idx + 1}
                                    </span>
                                    <span className="text-slate-200 font-medium group-hover:text-white transition-colors">
                                        {sheet}
                                    </span>
                                </div>
                                <Check className="w-4 h-4 text-emerald-500 opacity-0 group-hover:opacity-100 transition-opacity" />
                            </button>
                        ))}
                    </div>
                </div>
            </motion.div>
        </div>
    );
}
