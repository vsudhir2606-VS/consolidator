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
  const [consolidatedData, setConsolidatedData] = useState<any[] | null>(null);
  const [summary, setSummary] = useState<{ totalRows: number; totalFiles: number } | null>(null);

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files).map(file => ({
        file,
        id: Math.random().toString(36).substring(7),
        status: 'pending' as const
      }));
      setFiles(prev => [...prev, ...newFiles]);
      setConsolidatedData(null);
      setSummary(null);
      setError(null);
    }
  };

  const removeFile = (id: string) => {
    setFiles(prev => prev.filter(f => f.id !== id));
    setConsolidatedData(null);
    setSummary(null);
  };

  const consolidateFiles = async () => {
    if (files.length === 0) return;
    
    setIsConsolidating(true);
    setError(null);
    setConsolidatedData(null);

    try {
      let allRows: any[] = [];
      let processedFilesCount = 0;
      
      for (const fileStatus of files) {
        const data = await fileStatus.file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: false, cellText: false });
        
        let fileRowsCount = 0;
        
        // Process all sheets in the workbook
        for (const sheetName of workbook.SheetNames) {
          const worksheet = workbook.Sheets[sheetName];
          // Use header: 1 to get an array of arrays, including the first row
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" }) as any[][];
          
          if (jsonData.length > 0) {
            const dataWithSource = jsonData.map((row: any[]) => {
              const rowArray = Array.isArray(row) ? row : [row];
              return [...rowArray, fileStatus.file.name, sheetName];
            });
            
            // Use concat to avoid stack size limits with push(...spread)
            allRows = allRows.concat(dataWithSource);
            fileRowsCount += jsonData.length;
          }
        }
        
        processedFilesCount++;
        
        // Update status locally
        setFiles(prev => prev.map(f => 
          f.id === fileStatus.id ? { ...f, status: 'completed', rowCount: fileRowsCount } : f
        ));
      }

      if (allRows.length === 0) {
        throw new Error("No data found in any of the uploaded files/sheets.");
      }

      setConsolidatedData(allRows);
      setSummary({
        totalRows: allRows.length,
        totalFiles: processedFilesCount
      });

    } catch (err) {
      console.error("Consolidation Error:", err);
      setError(err instanceof Error ? err.message : "An error occurred during consolidation.");
    } finally {
      setIsConsolidating(false);
    }
  };

  const downloadConsolidated = () => {
    if (!consolidatedData) return;

    try {
      // Use aoa_to_sheet since we are now using array of arrays
      const newSheet = XLSX.utils.aoa_to_sheet(consolidatedData);
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Consolidated Data");

      const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      
      const link = document.createElement('a');
      link.href = url;
      link.download = `consolidated_report_${new Date().getTime()}.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (err) {
      setError("Failed to generate download file.");
    }
  };

  return (
    <div className="min-h-screen bg-zinc-50 p-6 md:p-12">
      <div className="max-w-5xl mx-auto">
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
            <h1 className="text-3xl font-bold tracking-tight text-zinc-900">Excel Master Consolidator</h1>
          </motion.div>
          <p className="text-zinc-500 text-lg max-w-2xl">
            Upload multiple files and we'll stack them vertically into one single sheet, preserving all columns and adding source metadata.
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

          <div className="grid lg:grid-cols-2 gap-8 items-start">
            {/* File List */}
            <AnimatePresence>
              {files.length > 0 && (
                <motion.section 
                  initial={{ opacity: 0, x: -20 }}
                  animate={{ opacity: 1, x: 0 }}
                  className="glass-card p-6"
                >
                  <div className="flex items-center justify-between mb-6">
                    <h2 className="text-lg font-bold flex items-center gap-2">
                      Queue <span className="text-sm font-normal text-zinc-400 bg-zinc-100 px-2 py-0.5 rounded-full">{files.length} files</span>
                    </h2>
                    <button 
                      onClick={() => { setFiles([]); setConsolidatedData(null); setSummary(null); }}
                      className="text-sm text-zinc-400 hover:text-red-500 transition-colors"
                    >
                      Clear all
                    </button>
                  </div>

                  <div className="space-y-2 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
                    {files.map((f) => (
                      <div 
                        key={f.id}
                        className="flex items-center justify-between p-3 bg-zinc-50 rounded-xl border border-zinc-200 group"
                      >
                        <div className="flex items-center gap-3 overflow-hidden">
                          <div className="p-2 bg-white rounded-lg border border-zinc-200 text-emerald-600">
                            <FileSpreadsheet size={16} />
                          </div>
                          <div className="truncate">
                            <p className="text-sm font-medium text-zinc-800 truncate">{f.file.name}</p>
                            <p className="text-[10px] text-zinc-400 uppercase tracking-wider font-semibold">
                              {(f.file.size / 1024).toFixed(1)} KB {f.rowCount !== undefined && `• ${f.rowCount} rows`}
                            </p>
                          </div>
                        </div>
                        
                        <div className="flex items-center gap-2">
                          {f.status === 'completed' && <CheckCircle2 size={16} className="text-emerald-500" />}
                          <button 
                            onClick={() => removeFile(f.id)}
                            className="p-1 text-zinc-300 hover:text-red-500 hover:bg-red-50 rounded-md transition-all"
                          >
                            <X size={16} />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>

                  <div className="mt-6">
                    <button
                      onClick={consolidateFiles}
                      disabled={isConsolidating || files.length === 0}
                      className={cn(
                        "w-full py-4 rounded-2xl font-bold text-lg flex items-center justify-center gap-3 transition-all",
                        isConsolidating 
                          ? "bg-zinc-100 text-zinc-400 cursor-not-allowed" 
                          : "bg-zinc-900 text-white hover:bg-black active:scale-[0.98]"
                      )}
                    >
                      {isConsolidating ? (
                        <>
                          <Loader2 className="animate-spin" />
                          Processing...
                        </>
                      ) : (
                        "Process & Consolidate"
                      )}
                    </button>
                  </div>
                </motion.section>
              )}
            </AnimatePresence>

            {/* Results / Preview */}
            <AnimatePresence>
              {(summary || error) && (
                <motion.section 
                  initial={{ opacity: 0, x: 20 }}
                  animate={{ opacity: 1, x: 0 }}
                  className="glass-card p-6 border-emerald-100 bg-emerald-50/10"
                >
                  {error ? (
                    <div className="flex flex-col items-center text-center py-8">
                      <div className="w-12 h-12 bg-red-100 text-red-600 rounded-full flex items-center justify-center mb-4">
                        <AlertCircle size={24} />
                      </div>
                      <h3 className="text-lg font-bold text-zinc-900 mb-2">Consolidation Failed</h3>
                      <p className="text-zinc-500 text-sm mb-6">{error}</p>
                      <button 
                        onClick={() => setError(null)}
                        className="text-emerald-600 font-semibold hover:underline"
                      >
                        Try again
                      </button>
                    </div>
                  ) : (
                    <div className="flex flex-col h-full">
                      <div className="flex items-center gap-4 mb-8">
                        <div className="w-12 h-12 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center">
                          <CheckCircle2 size={24} />
                        </div>
                        <div>
                          <h3 className="text-lg font-bold text-zinc-900">Ready for Download</h3>
                          <p className="text-zinc-500 text-sm">Successfully merged {summary?.totalFiles} files.</p>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-4 mb-8">
                        <div className="p-4 bg-white rounded-2xl border border-emerald-100">
                          <p className="text-[10px] uppercase tracking-widest text-zinc-400 font-bold mb-1">Total Rows</p>
                          <p className="text-2xl font-bold text-emerald-700">{summary?.totalRows.toLocaleString()}</p>
                        </div>
                        <div className="p-4 bg-white rounded-2xl border border-emerald-100">
                          <p className="text-[10px] uppercase tracking-widest text-zinc-400 font-bold mb-1">Status</p>
                          <p className="text-2xl font-bold text-emerald-700">Success</p>
                        </div>
                      </div>

                      {consolidatedData && consolidatedData.length > 0 && (
                        <div className="mb-8">
                          <p className="text-xs font-bold text-zinc-400 uppercase tracking-widest mb-3">Data Preview (First 5 rows)</p>
                          <div className="bg-white rounded-xl border border-zinc-200 overflow-hidden">
                            <div className="overflow-x-auto">
                              <table className="w-full text-[10px] text-left">
                                <thead className="bg-zinc-50 border-bottom border-zinc-200">
                                  <tr>
                                    {Array.from({ length: Math.min(consolidatedData[0]?.length || 0, 4) }).map((_, idx) => (
                                      <th key={idx} className="px-3 py-2 font-bold text-zinc-500 truncate max-w-[80px]">Col {idx + 1}</th>
                                    ))}
                                  </tr>
                                </thead>
                                <tbody>
                                  {consolidatedData.slice(0, 5).map((row, i) => (
                                    <tr key={i} className="border-t border-zinc-100">
                                      {Array.isArray(row) ? row.slice(0, 4).map((val: any, j: number) => (
                                        <td key={j} className="px-3 py-2 text-zinc-600 truncate max-w-[80px]">{String(val)}</td>
                                      )) : (
                                        <td className="px-3 py-2 text-zinc-600 truncate">{String(row)}</td>
                                      )}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        </div>
                      )}

                      <button
                        onClick={downloadConsolidated}
                        className="mt-auto w-full py-4 bg-emerald-600 text-white rounded-2xl font-bold text-lg flex items-center justify-center gap-3 hover:bg-emerald-700 shadow-lg shadow-emerald-200 transition-all active:scale-[0.98]"
                      >
                        <Download size={22} />
                        Download Consolidated File
                      </button>
                    </div>
                  )}
                </motion.section>
              )}
            </AnimatePresence>
          </div>

          {/* Empty State Info */}
          {files.length === 0 && (
            <section className="grid md:grid-cols-3 gap-6 mt-8">
              {[
                { title: "Vertical Stacking", desc: "Files are appended one after another, creating a single long list." },
                { title: "Column Union", desc: "If files have different columns, we'll include all of them in the final sheet." },
                { title: "Multi-Sheet", desc: "We automatically look through every sheet in every file you upload." }
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
          <p>© {new Date().getFullYear()} Excel Master Consolidator • All processing is client-side for maximum security.</p>
        </footer>
      </div>
      
      <style>{`
        .custom-scrollbar::-webkit-scrollbar {
          width: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #e4e4e7;
          border-radius: 10px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #d4d4d8;
        }
      `}</style>
    </div>
  );
}
