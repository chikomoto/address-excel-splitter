import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Upload, Download, FileSpreadsheet, CheckCircle2, AlertCircle } from "lucide-react";
import { motion } from "framer-motion";

function cleanText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function parseSingaporeAddress(value) {
  const raw = cleanText(value);
  if (!raw) return { address: "", unitNumber: "", postalCode: "" };

  const postalMatch = raw.match(/(\d{6})\s*$/);
  const postalCode = postalMatch ? postalMatch[1] : "";
  let working = raw;

  if (postalCode) {
    working = working.replace(new RegExp("\\s*" + postalCode + "\\s*$"), "").trim();
  }

  working = working.replace(/\s*SINGAPORE\s*$/i, "").trim();

  const unitMatch = working.match(/(#[A-Z0-9\-\/]+)$/i);
  const unitNumber = unitMatch ? unitMatch[1] : "";
  const address = unitNumber ? working.slice(0, working.lastIndexOf(unitNumber)).trim() : working;

  return { address, unitNumber, postalCode };
}

function looksLikeAddress(value) {
  const text = cleanText(value).toUpperCase();
  if (!text) return false;
  return /SINGAPORE\s+\d{6}$/.test(text) || (/#\d{2}-\d+/i.test(text) && /\d{6}$/.test(text));
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 30); i += 1) {
    const row = (rows[i] || []).map((cell) => cleanText(cell).toUpperCase());
    if (row.includes("PREMISES") || row.includes("ADDRESS")) return i;
  }
  return 0;
}

function findAddressColumn(headerRow, dataRows) {
  const headers = (headerRow || []).map((cell) => cleanText(cell).toUpperCase());
  const preferred = ["PREMISES", "ADDRESS", "FULL ADDRESS", "PREMISE ADDRESS", "PREMISES ADDRESS", "LOCATION", "SITE ADDRESS"];

  for (const name of preferred) {
    const idx = headers.findIndex((cell) => cell === name);
    if (idx !== -1) return idx;
  }

  let bestIndex = -1;
  let bestScore = 0;
  const maxColumns = Math.max(headerRow?.length || 0, ...dataRows.slice(0, 20).map((row) => row.length || 0));

  for (let col = 0; col < maxColumns; col += 1) {
    let score = 0;
    for (let row = 0; row < Math.min(dataRows.length, 20); row += 1) {
      if (looksLikeAddress(dataRows[row]?.[col])) score += 1;
    }
    if (score > bestScore) {
      bestScore = score;
      bestIndex = col;
    }
  }

  return bestScore > 0 ? bestIndex : -1;
}

function getColumnIndexFromLetter(letter) {
  const cleaned = String(letter || "").trim().toUpperCase();
  if (!cleaned) return 0;
  let result = 0;
  for (let i = 0; i < cleaned.length; i += 1) {
    result = result * 26 + (cleaned.charCodeAt(i) - 64);
  }
  return Math.max(result - 1, 0);
}

function isFooterRow(row) {
  const combined = row.map((cell) => cleanText(cell).toUpperCase()).join(" ");
  return combined.includes("-END OF REPORT-") || combined.includes("NO.OF CTR") || combined.includes("NO. OF CTR");
}

function buildProcessedSheet(rows, mode, headerName, columnLetter) {
  const headerRowIndex = findHeaderRow(rows);
  const headerRow = rows[headerRowIndex] || [];
  const dataRows = rows.slice(headerRowIndex + 1);

  let addressColIndex = 0;
  if (mode === "smart") {
    addressColIndex = findAddressColumn(headerRow, dataRows);
    if (addressColIndex === -1) throw new Error("Could not detect the address column automatically.");
  } else if (mode === "header") {
    addressColIndex = headerRow.findIndex((cell) => cleanText(cell).toLowerCase() === cleanText(headerName).toLowerCase());
    if (addressColIndex === -1) throw new Error("Column header \"" + headerName + "\" was not found.");
  } else {
    addressColIndex = getColumnIndexFromLetter(columnLetter);
  }

  const outputRows = rows.map((row) => [...row]);
  const originalHeader = cleanText(outputRows[headerRowIndex]?.[addressColIndex]) || "Address";
  outputRows[headerRowIndex].push(originalHeader + " - Parsed Address", originalHeader + " - Unit Number", originalHeader + " - Postal Code");

  let processedCount = 0;
  for (let i = headerRowIndex + 1; i < outputRows.length; i += 1) {
    const row = outputRows[i];
    if (isFooterRow(row)) {
      row.push("", "", "");
      continue;
    }
    const cellValue = row[addressColIndex];
    if (looksLikeAddress(cellValue)) {
      const parsed = parseSingaporeAddress(cellValue);
      row.push(parsed.address, parsed.unitNumber, parsed.postalCode);
      processedCount += 1;
    } else {
      row.push("", "", "");
    }
  }

  return {
    outputRows,
    previewRows: outputRows.slice(headerRowIndex, Math.min(outputRows.length, headerRowIndex + 8)),
    detectedHeaderRow: headerRowIndex,
    detectedColumnIndex: addressColIndex,
    detectedColumnName: originalHeader,
    processedCount,
  };
}

function App() {
  const [fileName, setFileName] = useState("");
  const [rowsPreview, setRowsPreview] = useState([]);
  const [status, setStatus] = useState("idle");
  const [message, setMessage] = useState("Upload an Excel file. Smart detect is tuned for SP-style reports with a PREMISES column.");
  const [processedWorkbook, setProcessedWorkbook] = useState(null);
  const [mode, setMode] = useState("smart");
  const [headerName, setHeaderName] = useState("PREMISES");
  const [columnLetter, setColumnLetter] = useState("F");
  const [sheetNameUsed, setSheetNameUsed] = useState("");
  const [detectedInfo, setDetectedInfo] = useState(null);

  const canDownload = useMemo(() => !!processedWorkbook, [processedWorkbook]);

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setStatus("processing");
      setMessage("Reading workbook...");
      setFileName(file.name);

      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetName = workbook.SheetNames?.[0];
      if (!sheetName) throw new Error("No worksheet found in the uploaded file.");

      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
      if (!rows.length) throw new Error("The selected sheet is empty.");

      const result = buildProcessedSheet(rows, mode, headerName, columnLetter);
      const newWorksheet = XLSX.utils.aoa_to_sheet(result.outputRows);
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

      setProcessedWorkbook(newWorkbook);
      setRowsPreview(result.previewRows);
      setSheetNameUsed(sheetName);
      setDetectedInfo({
        headerRowNumber: result.detectedHeaderRow + 1,
        columnNumber: result.detectedColumnIndex + 1,
        columnName: result.detectedColumnName,
        processedCount: result.processedCount,
      });
      setStatus("ready");
      setMessage("Done. Detected sheet \"" + sheetName + "\", header row " + (result.detectedHeaderRow + 1) + ", and address column \"" + result.detectedColumnName + "\".");
    } catch (error) {
      console.error(error);
      setStatus("error");
      setMessage(error.message || "Something went wrong while processing the Excel file.");
      setProcessedWorkbook(null);
      setRowsPreview([]);
      setDetectedInfo(null);
    }
  };

  const handleDownload = () => {
    if (!processedWorkbook) return;
    const outName = fileName.replace(/\.(xlsx|xls|csv)$/i, "") || "processed_addresses";
    XLSX.writeFile(processedWorkbook, outName + "_formatted.xlsx");
  };

  return (
    <div className="min-h-screen bg-[radial-gradient(circle_at_top,_rgba(139,92,246,0.18),_transparent_32%),radial-gradient(circle_at_bottom_right,_rgba(34,211,238,0.12),_transparent_28%),linear-gradient(135deg,#09090f_0%,#141126_45%,#0a0d18_100%)] text-slate-100 p-6 md:p-10">
      <div className="max-w-6xl mx-auto grid gap-6">
        <motion.div
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          className="relative overflow-hidden rounded-[28px] border border-violet-400/20 bg-white/5 shadow-[0_0_40px_rgba(139,92,246,0.18)] backdrop-blur-xl p-6 md:p-8"
        >
          <div className="absolute inset-0 bg-[linear-gradient(120deg,rgba(168,85,247,0.12),transparent_35%,rgba(34,211,238,0.08))] pointer-events-none" />
          <div className="relative">
            <div className="inline-flex items-center gap-2 rounded-full border border-violet-400/30 bg-violet-500/10 px-3 py-1 text-sm font-medium text-violet-200">
              <FileSpreadsheet className="h-4 w-4" />
              Geneco Address Splitter
            </div>
            <div className="mt-5 flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
              <div>
                <h1 className="text-3xl md:text-4xl font-semibold tracking-tight text-white">
                  Upload, split, and export Singapore addresses
                </h1>
                <p className="mt-3 max-w-3xl text-slate-300 leading-6">
                  Built for SP-style reports. The app auto-detects the real header row, finds the premises column, and appends separated address fields into a cleaned output file.
                </p>
              </div>
              <div className="grid grid-cols-2 gap-3 text-sm min-w-[260px]">
                <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3">
                  <div className="text-slate-400">Mode</div>
                  <div className="mt-1 font-medium text-white">{mode === "smart" ? "Smart Detect" : mode === "header" ? "Header Name" : "Column Letter"}</div>
                </div>
                <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3">
                  <div className="text-slate-400">Status</div>
                  <div className="mt-1 font-medium text-white capitalize">{status}</div>
                </div>
              </div>
            </div>
          </div>
        </motion.div>

        <div className="grid lg:grid-cols-[1.08fr_0.92fr] gap-6">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.05 }}
            className="rounded-[28px] border border-violet-400/20 bg-white/5 shadow-[0_0_30px_rgba(139,92,246,0.12)] backdrop-blur-xl p-6"
          >
            <div className="flex items-center justify-between gap-3 flex-wrap">
              <h2 className="text-xl font-semibold text-white">Upload settings</h2>
              <span className="rounded-full border border-cyan-400/25 bg-cyan-400/10 px-3 py-1 text-xs font-medium text-cyan-200">
                Excel in, Excel out
              </span>
            </div>

            <div className="mt-6 grid gap-5">
              <div>
                <label className="text-sm font-medium text-slate-300">Detection mode</label>
                <div className="mt-3 grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <button
                    onClick={() => setMode("smart")}
                    className={`rounded-2xl px-4 py-3 border transition text-sm font-medium ${
                      mode === "smart"
                        ? "border-violet-400/60 bg-gradient-to-r from-violet-600 to-fuchsia-600 text-white shadow-[0_0_20px_rgba(168,85,247,0.35)]"
                        : "border-white/10 bg-white/5 text-slate-200 hover:bg-white/10"
                    }`}
                  >
                    Smart detect
                  </button>
                  <button
                    onClick={() => setMode("header")}
                    className={`rounded-2xl px-4 py-3 border transition text-sm font-medium ${
                      mode === "header"
                        ? "border-violet-400/60 bg-gradient-to-r from-violet-600 to-fuchsia-600 text-white shadow-[0_0_20px_rgba(168,85,247,0.35)]"
                        : "border-white/10 bg-white/5 text-slate-200 hover:bg-white/10"
                    }`}
                  >
                    Use header name
                  </button>
                  <button
                    onClick={() => setMode("letter")}
                    className={`rounded-2xl px-4 py-3 border transition text-sm font-medium ${
                      mode === "letter"
                        ? "border-violet-400/60 bg-gradient-to-r from-violet-600 to-fuchsia-600 text-white shadow-[0_0_20px_rgba(168,85,247,0.35)]"
                        : "border-white/10 bg-white/5 text-slate-200 hover:bg-white/10"
                    }`}
                  >
                    Use column letter
                  </button>
                </div>
              </div>

              {mode === "header" && (
                <div>
                  <label className="text-sm font-medium text-slate-300">Address column header</label>
                  <input
                    value={headerName}
                    onChange={(e) => setHeaderName(e.target.value)}
                    className="mt-2 w-full rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-slate-100 outline-none focus:ring-2 focus:ring-violet-400/40"
                    placeholder="Example: PREMISES"
                  />
                </div>
              )}

              {mode === "letter" && (
                <div>
                  <label className="text-sm font-medium text-slate-300">Address column letter</label>
                  <input
                    value={columnLetter}
                    onChange={(e) => setColumnLetter(e.target.value.toUpperCase())}
                    className="mt-2 w-full rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-slate-100 outline-none focus:ring-2 focus:ring-violet-400/40"
                    placeholder="Example: F"
                  />
                </div>
              )}

              <div>
                <label className="text-sm font-medium text-slate-300">Excel file</label>
                <label className="group mt-3 relative flex cursor-pointer flex-col items-center justify-center gap-3 rounded-[24px] border-2 border-dashed border-violet-400/30 bg-white/5 px-6 py-12 text-center transition hover:bg-white/10 hover:border-violet-300/45 overflow-hidden">
                  <div className="absolute inset-0 bg-[radial-gradient(circle_at_top,_rgba(168,85,247,0.16),_transparent_38%)] opacity-80 pointer-events-none" />
                  <div className="relative flex h-14 w-14 items-center justify-center rounded-2xl bg-gradient-to-br from-violet-500 to-cyan-500 shadow-[0_0_25px_rgba(56,189,248,0.25)]">
                    <Upload className="h-6 w-6 text-white" />
                  </div>
                  <div className="relative">
                    <div className="font-medium text-white">Upload your report</div>
                    <div className="mt-1 text-sm text-slate-300">Choose .xlsx, .xls or .csv</div>
                  </div>
                  <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFileUpload} />
                </label>
                {fileName ? <p className="mt-3 text-sm text-slate-400">Selected file: {fileName}</p> : null}
              </div>

              <div className={`rounded-[22px] border p-4 ${
                status === "error"
                  ? "border-red-400/30 bg-red-500/10 text-red-200"
                  : status === "ready"
                  ? "border-emerald-400/30 bg-emerald-500/10 text-emerald-200"
                  : "border-white/10 bg-white/5 text-slate-300"
              }`}>
                <div className="flex items-center gap-2 font-medium">
                  {status === "error" ? <AlertCircle className="h-4 w-4" /> : <CheckCircle2 className="h-4 w-4" />}
                  Processing status
                </div>
                <p className="mt-2 text-sm leading-6">{message}</p>
              </div>

              {detectedInfo && (
                <div className="grid sm:grid-cols-2 gap-3 text-sm">
                  <div className="rounded-2xl border border-white/10 bg-white/5 p-4">
                    <div className="text-slate-400">Sheet</div>
                    <div className="mt-1 font-medium text-white">{sheetNameUsed}</div>
                  </div>
                  <div className="rounded-2xl border border-white/10 bg-white/5 p-4">
                    <div className="text-slate-400">Detected header row</div>
                    <div className="mt-1 font-medium text-white">{detectedInfo.headerRowNumber}</div>
                  </div>
                  <div className="rounded-2xl border border-white/10 bg-white/5 p-4 sm:col-span-2">
                    <div className="text-slate-400">Detected address column</div>
                    <div className="mt-1 font-medium text-white">{detectedInfo.columnName} (column {detectedInfo.columnNumber})</div>
                  </div>
                  <div className="rounded-2xl border border-white/10 bg-white/5 p-4 sm:col-span-2">
                    <div className="text-slate-400">Rows parsed</div>
                    <div className="mt-1 text-2xl font-semibold text-white">{detectedInfo.processedCount}</div>
                  </div>
                </div>
              )}

              <div className="flex gap-3 flex-wrap">
              <button
                onClick={handleDownload}
                disabled={!canDownload}
                className="inline-flex items-center justify-center gap-2 rounded-2xl bg-gradient-to-r from-violet-600 via-fuchsia-600 to-cyan-500 px-5 py-3 text-white font-medium shadow-[0_0_24px_rgba(168,85,247,0.35)] disabled:opacity-40 disabled:cursor-not-allowed"
              >
                <Download className="h-4 w-4" />
                Download updated file
              </button>

              <button
                onClick={() => {
                  setRowsPreview([]);
                  setDetectedInfo(null);
                  setFileName("");
                  setStatus("idle");
                  setMessage("Upload a file to begin.");
                }}
                className="inline-flex items-center justify-center gap-2 rounded-2xl border border-red-400/40 bg-red-500/10 px-5 py-3 text-red-200 hover:bg-red-500/20"
              >
                Clear / Start Over
              </button>
            </div>
            </div>
          </motion.div>

          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.1 }}
            className="rounded-[28px] border border-violet-400/20 bg-white/5 shadow-[0_0_30px_rgba(139,92,246,0.12)] backdrop-blur-xl p-6"
          >
            <div className="flex items-center justify-between gap-3 flex-wrap">
              <h2 className="text-xl font-semibold text-white">Output preview</h2>
              <span className="rounded-full border border-violet-400/25 bg-violet-500/10 px-3 py-1 text-xs font-medium text-violet-200">
                Live preview
              </span>
            </div>
            <p className="mt-2 text-sm text-slate-300 leading-6">
              The report layout is preserved. Three new columns are appended on the right side of the detected sheet.
            </p>

            <div className="mt-5 overflow-auto rounded-[22px] border border-white/10 bg-black/10">
              <table className="min-w-full text-sm border-collapse">
                <tbody>
                  {rowsPreview.length ? (
                    rowsPreview.map((row, rowIndex) => (
                      <tr key={rowIndex} className={rowIndex === 0 ? "bg-white/15 font-semibold sticky top-0" : "border-t border-white/10 odd:bg-white/5"}>
                        {row.map((cell, cellIndex) => (
                          <td key={cellIndex} className="px-4 py-2 whitespace-nowrap align-top text-slate-200 border-r border-white/10 last:border-r-0">
                            {String(cell ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td className="px-4 py-10 text-slate-400">
                        No preview yet. Upload a file to see the processed rows around the detected header.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            <div className="mt-5 grid gap-3">
              <div className="rounded-[22px] border border-white/10 bg-white/5 p-4">
                <p className="font-medium text-white">Expected parsing example</p>
                <p className="mt-2 text-sm text-slate-300">
                  Input: <span className="font-mono text-cyan-200">234 LOR 8 TOA PAYOH #01-284 SINGAPORE 310234</span>
                </p>
              </div>
              <div className="grid md:grid-cols-3 gap-3 text-sm">
                <div className="rounded-[22px] border border-white/10 bg-white/5 p-4">
                  <div className="text-slate-400">Parsed Address</div>
                  <div className="mt-2 font-medium text-white">234 LOR 8 TOA PAYOH</div>
                </div>
                <div className="rounded-[22px] border border-white/10 bg-white/5 p-4">
                  <div className="text-slate-400">Unit Number</div>
                  <div className="mt-2 font-medium text-white">#01-284</div>
                </div>
                <div className="rounded-[22px] border border-white/10 bg-white/5 p-4">
                  <div className="text-slate-400">Postal Code</div>
                  <div className="mt-2 font-medium text-white">310234</div>
                </div>
              </div>
            </div>
          </motion.div>
        </div>
      </div>
    </div>
  );
}

export default App;
