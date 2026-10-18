import React, { useState } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import {
    Wand2,
    Plus,
    Trash2,
    Play,
    Filter,
    Columns,
    Type,
    X
} from 'lucide-react';
import { PROCESSOR_TYPES, OPERATORS } from '../lib/processors';

export function AutomationPanel({ headers, rules, onRulesChange, isOpen, onClose }) {
    const [activeTab, setActiveTab] = useState('rules'); // 'rules' | 'history'

    const addRule = (type) => {
        const newRule = {
            id: Math.random().toString(36).substr(2, 9),
            type,
            columnName: headers[0], // Default to first column
            columnIndex: 0,
            enabled: true,
            // Default values
            operator: OPERATORS.CONTAINS,
            value: '',
            transformation: 'uppercase'
        };
        onRulesChange([...rules, newRule]);
    };

    const updateRule = (id, updates) => {
        onRulesChange(rules.map(r => r.id === id ? { ...r, ...updates } : r));
    };

    const removeRule = (id) => {
        onRulesChange(rules.filter(r => r.id !== id));
    };

    return (
        <AnimatePresence>
            {isOpen && (
                <>
                    {/* Backdrop */}
                    <motion.div
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        exit={{ opacity: 0 }}
                        onClick={onClose}
                        className="fixed inset-0 bg-black/40 backdrop-blur-sm z-40"
                    />

                    {/* Panel */}
                    <motion.div
                        initial={{ x: '100%' }}
                        animate={{ x: 0 }}
                        exit={{ x: '100%' }}
                        transition={{ type: 'spring', damping: 25, stiffness: 200 }}
                        className="fixed right-0 top-0 h-full w-full max-w-md bg-slate-900/90 border-l border-white/10 backdrop-blur-xl z-50 shadow-2xl overflow-hidden flex flex-col"
                    >
                        <div className="p-6 border-b border-white/10 flex items-center justify-between">
                            <div>
                                <h2 className="text-xl font-bold text-white flex items-center gap-2">
                                    <Wand2 className="w-5 h-5 text-[#818cf8]" />
                                    Automation
                                </h2>
                                <p className="text-sm text-slate-400">Build your processing workflow</p>
                            </div>
                            <button
                                onClick={onClose}
                                className="p-2 hover:bg-white/5 rounded-full text-slate-400 hover:text-white transition-colors"
                            >
                                <X className="w-5 h-5" />
                            </button>
                        </div>

                        <div className="flex-1 overflow-y-auto p-6 space-y-6">

                            {/* Quick Actions */}
                            <div className="grid grid-cols-3 gap-2">
                                <button
                                    onClick={() => addRule(PROCESSOR_TYPES.FILTER)}
                                    className="p-3 rounded-xl bg-slate-800/50 hover:bg-[#818cf8]/20 border border-white/5 hover:border-[#818cf8]/50 transition-all flex flex-col items-center gap-2 group"
                                >
                                    <Filter className="w-5 h-5 text-slate-400 group-hover:text-[#818cf8]" />
                                    <span className="text-xs text-slate-300 font-medium">Filter</span>
                                </button>
                                <button
                                    onClick={() => addRule(PROCESSOR_TYPES.TRANSFORM)}
                                    className="p-3 rounded-xl bg-slate-800/50 hover:bg-[#818cf8]/20 border border-white/5 hover:border-[#818cf8]/50 transition-all flex flex-col items-center gap-2 group"
                                >
                                    <Type className="w-5 h-5 text-slate-400 group-hover:text-[#818cf8]" />
                                    <span className="text-xs text-slate-300 font-medium">Transform</span>
                                </button>
                                <button
                                    onClick={() => addRule(PROCESSOR_TYPES.DELETE_COLUMN)}
                                    className="p-3 rounded-xl bg-slate-800/50 hover:bg-red-500/20 border border-white/5 hover:border-red-500/50 transition-all flex flex-col items-center gap-2 group"
                                >
                                    <Columns className="w-5 h-5 text-slate-400 group-hover:text-red-400" />
                                    <span className="text-xs text-slate-300 font-medium">Delete Col</span>
                                </button>
                            </div>

                            {/* Rules List */}
                            <div className="space-y-4">
                                <h3 className="text-sm font-semibold text-slate-400 uppercase tracking-wider">Active Rules</h3>

                                <AnimatePresence mode="popLayout">
                                    {rules.length === 0 ? (
                                        <motion.div
                                            key="empty"
                                            initial={{ opacity: 0 }}
                                            animate={{ opacity: 1 }}
                                            className="text-center py-8 border-2 border-dashed border-slate-800 rounded-xl"
                                        >
                                            <p className="text-slate-500 text-sm">No rules added yet.</p>
                                            <p className="text-slate-600 text-xs mt-1">Click a quick action above to start.</p>
                                        </motion.div>
                                    ) : (
                                        rules.map((rule, idx) => (
                                            <motion.div
                                                layout
                                                key={rule.id}
                                                initial={{ opacity: 0, x: 20 }}
                                                animate={{ opacity: 1, x: 0 }}
                                                exit={{ opacity: 0, x: -20 }}
                                                className="bg-slate-800/40 rounded-xl border border-white/5 p-4 relative group hover:border-white/10 transition-colors"
                                            >
                                                <div className="absolute right-2 top-2 opacity-0 group-hover:opacity-100 transition-opacity">
                                                    <button
                                                        onClick={() => removeRule(rule.id)}
                                                        className="p-1.5 hover:bg-red-500/20 rounded-lg text-slate-500 hover:text-red-400"
                                                    >
                                                        <Trash2 className="w-4 h-4" />
                                                    </button>
                                                </div>

                                                <div className="flex items-center gap-2 mb-3">
                                                    <span className="text-xs font-mono text-slate-500 bg-slate-900 rounded px-1.5 py-0.5">#{idx + 1}</span>
                                                    <span className="text-sm font-medium text-[#818cf8] capitalize">
                                                        {rule.type.replace('_', ' ')}
                                                    </span>
                                                </div>

                                                <div className="space-y-3">
                                                    {/* Column Selector */}
                                                    <div className="flex flex-col gap-1">
                                                        <label className="text-[10px] text-slate-500 uppercase font-semibold">Column</label>
                                                        <select
                                                            value={rule.columnName}
                                                            onChange={(e) => updateRule(rule.id, {
                                                                columnName: e.target.value,
                                                                columnIndex: headers.indexOf(e.target.value)
                                                            })}
                                                            className="bg-slate-900/50 border border-slate-700 text-sm rounded-lg p-2 text-slate-200 outline-none focus:border-[#818cf8]"
                                                        >
                                                            {headers.map(h => (
                                                                <option key={h} value={h}>{h}</option>
                                                            ))}
                                                        </select>
                                                    </div>

                                                    {/* Type Specific Fields */}
                                                    {rule.type === PROCESSOR_TYPES.FILTER && (
                                                        <div className="grid grid-cols-2 gap-2">
                                                            <div className="flex flex-col gap-1">
                                                                <label className="text-[10px] text-slate-500 uppercase font-semibold">Operator</label>
                                                                <select
                                                                    value={rule.operator}
                                                                    onChange={(e) => updateRule(rule.id, { operator: e.target.value })}
                                                                    className="bg-slate-900/50 border border-slate-700 text-sm rounded-lg p-2 text-slate-200 outline-none focus:border-[#818cf8]"
                                                                >
                                                                    {Object.entries(OPERATORS).map(([k, v]) => (
                                                                        <option key={k} value={v}>{k.replace('_', ' ')}</option>
                                                                    ))}
                                                                </select>
                                                            </div>
                                                            <div className="flex flex-col gap-1">
                                                                <label className="text-[10px] text-slate-500 uppercase font-semibold">Value</label>
                                                                <input
                                                                    type="text"
                                                                    value={rule.value}
                                                                    onChange={(e) => updateRule(rule.id, { value: e.target.value })}
                                                                    className="bg-slate-900/50 border border-slate-700 text-sm rounded-lg p-2 text-slate-200 outline-none focus:border-[#818cf8]"
                                                                    placeholder="Match value..."
                                                                />
                                                            </div>
                                                        </div>
                                                    )}

                                                    {rule.type === PROCESSOR_TYPES.TRANSFORM && (
                                                        <div className="flex flex-col gap-1">
                                                            <label className="text-[10px] text-slate-500 uppercase font-semibold">Transformation</label>
                                                            <select
                                                                value={rule.transformation}
                                                                onChange={(e) => updateRule(rule.id, { transformation: e.target.value })}
                                                                className="bg-slate-900/50 border border-slate-700 text-sm rounded-lg p-2 text-slate-200 outline-none focus:border-[#818cf8]"
                                                            >
                                                                <option value="uppercase">UPPERCASE</option>
                                                                <option value="lowercase">lowercase</option>
                                                                <option value="trim">Trim Whitespace</option>
                                                            </select>
                                                        </div>
                                                    )}
                                                </div>
                                            </motion.div>
                                        ))
                                    )}
                                </AnimatePresence>
                            </div>
                        </div>

                        <div className="p-6 border-t border-white/10 bg-slate-900/50">
                            <button
                                onClick={onClose}
                                className="w-full btn flex items-center justify-center gap-2"
                            >
                                <Play className="w-4 h-4 fill-current" />
                                Done & Process
                            </button>
                        </div>
                    </motion.div>
                </>
            )}
        </AnimatePresence>
    );
}
