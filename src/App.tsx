/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  FileSpreadsheet, 
  X, 
  Download, 
  Loader2, 
  CheckCircle2, 
  AlertCircle,
  Plus
} from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { motion, AnimatePresence } from 'motion/react';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface FileStatus {
  file: File;
  id: string;
  status: 'pending' | 'processing' | 'completed' | 'error';
  rowCount?: number;
}

export default function App() {
  const [files, setFiles] = useState<FileStatus[]>([]);
  const [isConsolidating, setIsConsolidating] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files).map(file => ({
        file,
        id: Math.random().toString(36).substring(7),
        status: 'pending' as const
      }));
      setFiles(prev => [...prev, ...newFiles]);
      setError(null);
    }
  };

  const removeFile = (id: string) => {
    setFiles(prev => prev.filter(f => f.id !== id));
  };

  const consolidateFiles = async () => {
    if (files.length === 0) return;
    
    setIsConsolidating(true);
    setError(null);

    try {
      const allData: any[] = [];
      
      for (const fileStatus of files) {
        const data = await fileStatus.file.arrayBuffer();
        const workbook = XLSX.read(data);
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        // Add a source column to track where data came from
        const dataWithSource = jsonData.map((row: any) => ({
          ...row,
          '_source_file': fileStatus.file.name
        }));
        
        allData.push(...dataWithSource);
        
        // Update status locally
        setFiles(prev => prev.map(f => 
          f.id === fileStatus.id ? { ...f, status: 'completed', rowCount: jsonData.length } : f
        ));
      }

      if (allData.length === 0) {
        throw new Error("No data found in the uploaded files.");
      }

      // Create new workbook
      const newSheet = XLSX.utils.json_to_sheet(allData);
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Consolidated");

      // Generate buffer and download
      const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      
      const link = document.createElement('a');
      link.href = url;
      link.download = `consolidated_data_${new Date().toISOString().split('T')[0]}.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

    } catch (err) {
      console.error(err);
      setError(err instanceof Error ? err.message : "An error occurred during consolidation.");
    } finally {
      setIsConsolidating(false);
    }
  };

  return (
    <div className="min-h-screen bg-zinc-50 p-6 md:p-12">
      <div className="max-w-4xl mx-auto">
        {/* Header */}
        <header className="mb-12">
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            className="flex items-center gap-3 mb-4"
          >
            <div className="p-3 bg-emerald-600 rounded-xl text-white shadow-lg shadow-emerald-200">
              <FileSpreadsheet size={28} />
            </div>
            <h1 className="text-3xl font-bold tracking-tight text-zinc-900">Excel Consolidator</h1>
          </motion.div>
          <p className="text-zinc-500 text-lg max-w-2xl">
            Combine multiple Excel files into one master sheet. We'll even add a source column so you know where each row came from.
          </p>
        </header>

        <main className="grid gap-8">
          {/* Upload Area */}
          <section>
            <label 
              className={cn(
                "relative group cursor-pointer block",
                "border-2 border-dashed border-zinc-300 rounded-3xl p-12 transition-all",
                "hover:border-emerald-500 hover:bg-emerald-50/50",
                files.length > 0 ? "bg-white" : "bg-zinc-100/50"
              )}
            >
              <input 
                type="file" 
                multiple 
                accept=".xlsx, .xls, .csv" 
                className="hidden" 
                onChange={onFileChange}
              />
              <div className="flex flex-col items-center text-center">
                <div className="w-16 h-16 bg-zinc-200 rounded-full flex items-center justify-center mb-4 group-hover:bg-emerald-100 group-hover:text-emerald-600 transition-colors">
                  <Upload size={32} />
                </div>
                <h3 className="text-xl font-semibold text-zinc-800 mb-2">
                  {files.length > 0 ? "Add more files" : "Select Excel files to merge"}
                </h3>
                <p className="text-zinc-500">
                  Drag and drop files here or click to browse
                </p>
              </div>
            </label>
          </section>

          {/* File List */}
          <AnimatePresence>
            {files.length > 0 && (
              <motion.section 
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 0.95 }}
                className="glass-card p-6"
              >
                <div className="flex items-center justify-between mb-6">
                  <h2 className="text-lg font-bold flex items-center gap-2">
                    Files to process <span className="text-sm font-normal text-zinc-400 bg-zinc-100 px-2 py-0.5 rounded-full">{files.length}</span>
                  </h2>
                  <button 
                    onClick={() => setFiles([])}
                    className="text-sm text-zinc-400 hover:text-red-500 transition-colors"
                  >
                    Clear all
                  </button>
                </div>

                <div className="space-y-3">
                  {files.map((f) => (
                    <motion.div 
                      layout
                      initial={{ opacity: 0, x: -10 }}
                      animate={{ opacity: 1, x: 0 }}
                      key={f.id}
                      className="flex items-center justify-between p-4 bg-zinc-50 rounded-xl border border-zinc-200 group"
                    >
                      <div className="flex items-center gap-4 overflow-hidden">
                        <div className="p-2 bg-white rounded-lg border border-zinc-200 text-emerald-600">
                          <FileSpreadsheet size={20} />
                        </div>
                        <div className="truncate">
                          <p className="font-medium text-zinc-800 truncate">{f.file.name}</p>
                          <p className="text-xs text-zinc-400">
                            {(f.file.size / 1024).toFixed(1)} KB • {f.rowCount ? `${f.rowCount} rows` : 'Pending'}
                          </p>
                        </div>
                      </div>
                      
                      <div className="flex items-center gap-3">
                        {f.status === 'completed' && <CheckCircle2 size={18} className="text-emerald-500" />}
                        <button 
                          onClick={() => removeFile(f.id)}
                          className="p-1 text-zinc-300 hover:text-red-500 hover:bg-red-50 rounded-md transition-all"
                        >
                          <X size={18} />
                        </button>
                      </div>
                    </motion.div>
                  ))}
                </div>

                <div className="mt-8 flex flex-col gap-4">
                  {error && (
                    <div className="p-4 bg-red-50 border border-red-100 rounded-xl flex items-center gap-3 text-red-700 text-sm">
                      <AlertCircle size={18} />
                      {error}
                    </div>
                  )}

                  <button
                    onClick={consolidateFiles}
                    disabled={isConsolidating || files.length === 0}
                    className={cn(
                      "w-full py-4 rounded-2xl font-bold text-lg flex items-center justify-center gap-3 transition-all shadow-lg shadow-emerald-100",
                      isConsolidating 
                        ? "bg-zinc-200 text-zinc-400 cursor-not-allowed" 
                        : "bg-emerald-600 text-white hover:bg-emerald-700 active:scale-[0.98]"
                    )}
                  >
                    {isConsolidating ? (
                      <>
                        <Loader2 className="animate-spin" />
                        Consolidating...
                      </>
                    ) : (
                      <>
                        <Download size={22} />
                        Consolidate & Download
                      </>
                    )}
                  </button>
                </div>
              </motion.section>
            )}
          </AnimatePresence>

          {/* Empty State Info */}
          {files.length === 0 && (
            <section className="grid md:grid-cols-3 gap-6 mt-8">
              {[
                { title: "Multi-file Support", desc: "Upload as many Excel or CSV files as you need." },
                { title: "Smart Merging", desc: "We combine data based on column headers automatically." },
                { title: "Source Tracking", desc: "Every row gets a '_source_file' column for easy auditing." }
              ].map((item, i) => (
                <div key={i} className="p-6 bg-white border border-zinc-200 rounded-2xl">
                  <h4 className="font-bold text-zinc-800 mb-2">{item.title}</h4>
                  <p className="text-sm text-zinc-500 leading-relaxed">{item.desc}</p>
                </div>
              ))}
            </section>
          )}
        </main>

        <footer className="mt-20 text-center text-zinc-400 text-sm">
          <p>© {new Date().getFullYear()} Excel Consolidator • Privacy focused: All processing happens in your browser.</p>
        </footer>
      </div>
    </div>
  );
}
