import React, { useState, useMemo } from 'react';
import { motion } from 'framer-motion';
import {
    ChevronLeft,
    ChevronRight,
    Search,
    ArrowUpDown,
    Download,
    Filter
} from 'lucide-react';
import * as XLSX from 'xlsx';

export function DataTable({ data, searchTerm = '' }) {
    const [currentPage, setCurrentPage] = useState(1);
    const [rowsPerPage, setRowsPerPage] = useState(10);

    // Headers are the first row, data is the rest
    const { headers, rows } = useMemo(() => ({
        headers: data[0] || [],
        rows: data.slice(1) || []
    }), [data]);

    const filteredRows = useMemo(() => {
        if (!searchTerm) return rows;
        const lowerSearch = searchTerm.toLowerCase();
        return rows.filter(row =>
            row.some(cell =>
                cell !== null && cell !== undefined && String(cell).toLowerCase().includes(lowerSearch)
            )
        );
    }, [rows, searchTerm]);

    const totalPages = Math.ceil(filteredRows.length / rowsPerPage);

    // Reset page when search changes
    React.useEffect(() => {
        setCurrentPage(1);
    }, [searchTerm]);

    const currentRows = filteredRows.slice(
        (currentPage - 1) * rowsPerPage,
        currentPage * rowsPerPage
    );

    if (!data || data.length === 0) return null;

    return (
        <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="space-y-4 w-full"
        >
            <div className="glass-card overflow-hidden !p-0">
                <div className="overflow-x-auto">
                    <table className="w-full text-sm text-left text-slate-400">
                        <thead className="text-xs uppercase bg-slate-900/40 text-slate-300">
                            <tr>
                                <th className="px-6 py-4 w-16 text-center">#</th>
                                {headers.map((header, idx) => {
                                    // Align Logic: Center by default, Left for Name
                                    const isName = header && header.toString().toLowerCase().includes('name');
                                    const alignClass = isName ? 'text-left' : 'text-center';

                                    return (
                                        <th key={idx} className={`px-6 py-4 font-semibold whitespace-nowrap group cursor-pointer hover:bg-white/5 transition-colors ${alignClass}`}>
                                            <div className={`flex items-center gap-2 ${isName ? 'justify-start' : 'justify-center'}`}>
                                                {header || `Column ${idx + 1}`}
                                                <ArrowUpDown className="w-3 h-3 opacity-0 group-hover:opacity-50" />
                                            </div>
                                        </th>
                                    );
                                })}
                            </tr>
                        </thead>
                        <tbody>
                            {currentRows.map((row, rowIndex) => (
                                <tr
                                    key={rowIndex}
                                    className="border-b border-slate-700/50 hover:bg-white/5 transition-colors last:border-0 odd:bg-slate-800/10"
                                >
                                    <td className="px-6 py-4 text-center text-slate-500 font-mono text-xs">
                                        {((currentPage - 1) * rowsPerPage) + rowIndex + 1}
                                    </td>
                                    {row.map((cell, cellIndex) => {
                                        const header = headers[cellIndex];
                                        const isName = header && header.toString().toLowerCase().includes('name');
                                        const alignClass = isName ? 'text-left' : 'text-center';

                                        return (
                                            <td key={cellIndex} className={`px-6 py-4 text-slate-300 whitespace-nowrap max-w-xs overflow-hidden text-ellipsis ${alignClass}`}>
                                                {cell}
                                            </td>
                                        );
                                    })}
                                    {/* Fill empty cells if row length mismatch */}
                                    {Array.from({ length: Math.max(0, headers.length - row.length) }).map((_, i) => (
                                        <td key={`empty-${i}`} className="px-6 py-4"></td>
                                    ))}
                                </tr>
                            ))}
                            {currentRows.length === 0 && (
                                <tr>
                                    <td colSpan={headers.length + 1} className="px-6 py-12 text-center text-slate-500">
                                        No results found for "{searchTerm}"
                                    </td>
                                </tr>
                            )}
                        </tbody>
                    </table>
                </div>

                {/* Pagination */}
                <div className="flex items-center justify-between p-4 bg-slate-900/20 border-t border-slate-700/50">
                    <div className="flex items-center gap-2 text-sm text-slate-400">
                        <span>Rows per page:</span>
                        <select
                            value={rowsPerPage}
                            onChange={(e) => {
                                setRowsPerPage(Number(e.target.value));
                                setCurrentPage(1);
                            }}
                            className="bg-slate-800 border border-slate-700 rounded px-2 py-1 focus:outline-none focus:border-[#818cf8]"
                        >
                            <option value={10}>10</option>
                            <option value={20}>20</option>
                            <option value={50}>50</option>
                            <option value={100}>100</option>
                        </select>
                    </div>

                    <div className="flex items-center gap-2">
                        <button
                            onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                            disabled={currentPage === 1}
                            className="p-1 rounded hover:bg-white/10 disabled:opacity-30 disabled:hover:bg-transparent transition-colors"
                        >
                            <ChevronLeft className="w-5 h-5" />
                        </button>
                        <span className="text-sm text-slate-300">
                            Page <span className="font-semibold text-white">{currentPage}</span> of {totalPages || 1}
                        </span>
                        <button
                            onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                            disabled={currentPage === totalPages || totalPages === 0}
                            className="p-1 rounded hover:bg-white/10 disabled:opacity-30 disabled:hover:bg-transparent transition-colors"
                        >
                            <ChevronRight className="w-5 h-5" />
                        </button>
                    </div>
                </div>
            </div>
        </motion.div>
    );
}
