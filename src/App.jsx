import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  Download,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
} from "lucide-react";
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

  const unitMatch = working.match(/(#[A-Z0-9\-/]+)$/i);
  const unitNumber = unitMatch ? unitMatch[1] : "";
  const address = unitNumber
    ? working.slice(0, working.lastIndexOf(unitNumber)).trim()
    : working;

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
  const preferred = [
    "PREMISES",
    "ADDRESS",
    "FULL ADDRESS",
    "PREMISE ADDRESS",
    "PREMISES ADDRESS",
    "LOCATION",
    "SITE ADDRESS",
  ];

  for (const name of preferred) {
    const idx = headers.findIndex((cell) => cell === name);
    if (idx !== -1) return idx;
  }

  let bestIndex = -1;
  let bestScore = 0;
  const widthCandidates = [headerRow?.length || 0, ...dataRows.slice(0, 20).map((row) => row.length || 0)];
  const maxColumns = Math.max(...widthCandidates);

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
  return (
    combined.includes("-END OF REPORT-") ||
    combined.includes("NO.OF CTR") ||
    combined.includes("NO. OF CTR")
  );
}

function buildProcessedSheet(rows, mode, headerName, columnLetter) {
  const headerRowIndex = findHeaderRow(rows);
  const headerRow = rows[headerRowIndex] || [];
  const dataRows = rows.slice(headerRowIndex + 1);

  let addressColIndex = 0;
  if (mode === "smart") {
    addressColIndex = findAddressColumn(headerRow, dataRows);
    if (addressColIndex === -1) {
      throw new Error("Could not detect the address column automatically.");
    }
  } else if (mode === "header") {
    addressColIndex = headerRow.findIndex(
      (cell) => cleanText(cell).toLowerCase() === cleanText(headerName).toLowerCase()
    );
    if (addressColIndex === -1) {
      throw new Error(`Column header "${headerName}" was not found.`);
    }
  } else {
    addressColIndex = getColumnIndexFromLetter(columnLetter);
  }

  const outputRows = rows.map((row) => [...row]);
  const originalHeader = cleanText(outputRows[headerRowIndex]?.[addressColIndex]) || "Address";

  outputRows[headerRowIndex].push(
    originalHeader + " - Parsed Address",
    originalHeader + " - Unit Number",
    originalHeader + " - Postal Code"
  );

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
    previewRows: outputRows.slice(
      headerRowIndex,
      Math.min(outputRows.length, headerRowIndex + 8)
    ),
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
  const [message, setMessage] = useState(
    "Upload an Excel file. Smart detect is tuned for SP-style reports with a PREMISES column."
  );
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
      setMessage(
        `Done. Detected sheet "${sheetName}", header row ${
          result.detectedHeaderRow + 1
        }, and address column "${result.detectedColumnName}".`
      );
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
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-slate-100 text-slate-900 p-6 md:p-10">
      <div className="max-w-6xl mx-auto grid gap-6">
        <motion.div
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          className="rounded-3xl border border-slate-200 bg-white/90 shadow-xl p-6 md:p-8"
        >
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-slate-200 px-3 py-1 text-sm font-medium text-slate-600">
              <FileSpreadsheet className="h-4 w-4" />
              Excel Address Splitter
            </div>
            <h1 className="mt-4 text-3xl md:text-4xl font-semibold tracking-tight">
              Upload an Excel file and split Singapore addresses automatically
            </h1>
            <p className="mt-3 text-slate-600 max-w-3xl">
              Rewritten to support SP-style reports where the real header row is lower down
              and the address column is usually named PREMISES.
            </p>
          </div>
        </motion.div>

        <div className="grid lg:grid-cols-[1.1fr_0.9fr] gap-6">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.05 }}
            className="rounded-3xl border border-slate-200 bg-white shadow-lg p-6"
          >
            <h2 className="text-xl font-semibold">1. Configure input</h2>
            <div className="mt-5 grid gap-4">
              <div>
                <label className="text-sm font-medium text-slate-700">Detection mode</label>
                <div className="mt-2 flex gap-2 flex-wrap">
                  <button
                    onClick={() => setMode("smart")}
                    className={`rounded-2xl px-4 py-2 border transition ${
                      mode === "smart"
                        ? "border-slate-900 bg-slate-900 text-white"
                        : "border-slate-300 bg-white text-slate-700"
                    }`}
                  >
                    Smart detect
                  </button>
                  <button
                    onClick={() => setMode("header")}
                    className={`rounded-2xl px-4 py-2 border transition ${
                      mode === "header"
                        ? "border-slate-900 bg-slate-900 text-white"
                        : "border-slate-300 bg-white text-slate-700"
                    }`}
                  >
                    Use header name
                  </button>
                  <button
                    onClick={() => setMode("letter")}
                    className={`rounded-2xl px-4 py-2 border transition ${
                      mode === "letter"
                        ? "border-slate-900 bg-slate-900 text-white"
                        : "border-slate-300 bg-white text-slate-700"
                    }`}
                  >
                    Use column letter
                  </button>
                </div>
              </div>

              {mode === "header" && (
                <div>
                  <label className="text-sm font-medium text-slate-700">
                    Address column header
                  </label>
                  <input
                    value={headerName}
                    onChange={(e) => setHeaderName(e.target.value)}
                    className="mt-2 w-full rounded-2xl border border-slate-300 px-4 py-3 outline-none focus:ring-2 focus:ring-slate-300"
                    placeholder="Example: PREMISES"
                  />
                </div>
              )}

              {mode === "letter" && (
                <div>
                  <label className="text-sm font-medium text-slate-700">
                    Address column letter
                  </label>
                  <input
                    value={columnLetter}
                    onChange={(e) => setColumnLetter(e.target.value.toUpperCase())}
                    className="mt-2 w-full rounded-2xl border border-slate-300 px-4 py-3 outline-none focus:ring-2 focus:ring-slate-300"
                    placeholder="Example: F"
                  />
                </div>
              )}

              <div>
                <label className="text-sm font-medium text-slate-700">Excel file</label>
                <label className="mt-2 flex cursor-pointer items-center justify-center gap-3 rounded-3xl border-2 border-dashed border-slate-300 bg-slate-50 px-6 py-10 text-center hover:bg-slate-100 transition">
                  <Upload className="h-5 w-5" />
                  <span className="font-medium">Choose .xlsx, .xls or .csv file</span>
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    className="hidden"
                    onChange={handleFileUpload}
                  />
                </label>
                {fileName ? (
                  <p className="mt-2 text-sm text-slate-500">Selected file: {fileName}</p>
                ) : null}
              </div>

              <div
                className={`rounded-2xl border p-4 ${
                  status === "error"
                    ? "border-red-200 bg-red-50 text-red-700"
                    : status === "ready"
                    ? "border-emerald-200 bg-emerald-50 text-emerald-700"
                    : "border-slate-200 bg-slate-50 text-slate-600"
                }`}
              >
                <div className="flex items-center gap-2 font-medium">
                  {status === "error" ? (
                    <AlertCircle className="h-4 w-4" />
                  ) : (
                    <CheckCircle2 className="h-4 w-4" />
                  )}
                  Status
                </div>
                <p className="mt-1 text-sm">{message}</p>
              </div>

              {detectedInfo && (
                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4 text-sm text-slate-700">
                  <div><strong>Sheet:</strong> {sheetNameUsed}</div>
                  <div><strong>Detected header row:</strong> {detectedInfo.headerRowNumber}</div>
                  <div>
                    <strong>Detected address column:</strong> {detectedInfo.columnName} (column{" "}
                    {detectedInfo.columnNumber})
                  </div>
                  <div><strong>Rows parsed:</strong> {detectedInfo.processedCount}</div>
                </div>
              )}

              <button
                onClick={handleDownload}
                disabled={!canDownload}
                className="inline-flex items-center justify-center gap-2 rounded-2xl bg-slate-900 px-5 py-3 text-white font-medium disabled:opacity-40 disabled:cursor-not-allowed"
              >
                <Download className="h-4 w-4" />
                Download updated file
              </button>
            </div>
          </motion.div>

          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.1 }}
            className="rounded-3xl border border-slate-200 bg-white shadow-lg p-6"
          >
            <h2 className="text-xl font-semibold">2. Preview output</h2>
            <p className="mt-2 text-sm text-slate-600">
              The app keeps the report layout and appends 3 new columns to the right side
              of the detected sheet.
            </p>

            <div className="mt-5 overflow-auto rounded-2xl border border-slate-200">
              <table className="min-w-full text-sm">
                <tbody>
                  {rowsPreview.length ? (
                    rowsPreview.map((row, rowIndex) => (
                      <tr
                        key={rowIndex}
                        className={
                          rowIndex === 0 ? "bg-slate-100 font-semibold" : "border-t border-slate-200"
                        }
                      >
                        {row.map((cell, cellIndex) => (
                          <td key={cellIndex} className="px-3 py-2 whitespace-nowrap align-top">
                            {String(cell ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td className="px-4 py-8 text-slate-500">
                        No preview yet. Upload a file to see the processed rows around the
                        detected header.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </motion.div>
        </div>
      </div>
    </div>
  );
}

export default App;