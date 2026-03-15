import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Upload, Download, FileSpreadsheet, CheckCircle2, AlertCircle } from "lucide-react";
import { motion } from "framer-motion";

function parseSingaporeAddress(value) {
  const raw = String(value ?? "").trim().replace(/\s+/g, " ");
  if (!raw) {
    return {
      address: "",
      unitNumber: "",
      postalCode: "",
    };
  }

  const postalMatch = raw.match(/(\d{6})\s*$/);
  const postalCode = postalMatch ? postalMatch[1] : "";

  let working = raw;
  if (postalCode) {
    working = working.replace(new RegExp(`\\s*${postalCode}\\s*$`), "").trim();
  }

  working = working.replace(/\s*SINGAPORE\s*$/i, "").trim();

  const unitMatch = working.match(/(#[A-Z0-9-\/]+)$/i);
  const unitNumber = unitMatch ? unitMatch[1] : "";

  let address = working;
  if (unitNumber) {
    address = working.slice(0, working.lastIndexOf(unitNumber)).trim();
  }

  return {
    address,
    unitNumber,
    postalCode,
  };
}

function App() {
  const [fileName, setFileName] = useState("");
  const [rowsPreview, setRowsPreview] = useState([]);
  const [status, setStatus] = useState("idle");
  const [message, setMessage] = useState("Upload an Excel file that contains a full address column.");
  const [processedWorkbook, setProcessedWorkbook] = useState(null);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [availableSheets, setAvailableSheets] = useState([]);
  const [columnMode, setColumnMode] = useState("header");
  const [headerName, setHeaderName] = useState("Address");
  const [columnLetter, setColumnLetter] = useState("A");

  const canDownload = useMemo(() => !!processedWorkbook, [processedWorkbook]);

  const getColumnIndexFromLetter = (letter) => {
    const cleaned = String(letter || "").trim().toUpperCase();
    if (!cleaned) return 0;
    let result = 0;
    for (let i = 0; i < cleaned.length; i += 1) {
      result = result * 26 + (cleaned.charCodeAt(i) - 64);
    }
    return Math.max(result - 1, 0);
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setStatus("processing");
      setMessage("Reading workbook...");
      setFileName(file.name);

      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetNames = workbook.SheetNames || [];
      setAvailableSheets(sheetNames);
      const sheetName = selectedSheet || sheetNames[0];
      setSelectedSheet(sheetName);

      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

      if (!rows.length) {
        setStatus("error");
        setMessage("The selected sheet is empty.");
        setProcessedWorkbook(null);
        return;
      }

      let addressColIndex = 0;
      let outputRows = [];

      if (columnMode === "header") {
        const headerRow = rows[0].map((v) => String(v ?? "").trim());
        addressColIndex = headerRow.findIndex(
          (v) => v.toLowerCase() === headerName.trim().toLowerCase()
        );

        if (addressColIndex === -1) {
          setStatus("error");
          setMessage(`Column header \"${headerName}\" was not found in sheet \"${sheetName}\".`);
          setProcessedWorkbook(null);
          return;
        }

        outputRows = [
          [...headerRow, "Parsed Address", "Unit Number", "Postal Code"],
          ...rows.slice(1).map((row) => {
            const parsed = parseSingaporeAddress(row[addressColIndex]);
            return [...row, parsed.address, parsed.unitNumber, parsed.postalCode];
          }),
        ];
      } else {
        addressColIndex = getColumnIndexFromLetter(columnLetter);
        outputRows = rows.map((row, index) => {
          const parsed = index === 0 && !String(row[addressColIndex] ?? "").includes("#") && !String(row[addressColIndex] ?? "").match(/\d{6}$/)
            ? { address: "", unitNumber: "", postalCode: "" }
            : parseSingaporeAddress(row[addressColIndex]);
          return [...row, parsed.address, parsed.unitNumber, parsed.postalCode];
        });

        if (outputRows.length) {
          outputRows[0] = [...outputRows[0], "Parsed Address", "Unit Number", "Postal Code"];
        }
      }

      const newWorksheet = XLSX.utils.aoa_to_sheet(outputRows);
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

      setProcessedWorkbook(newWorkbook);
      setRowsPreview(outputRows.slice(0, 8));
      setStatus("ready");
      setMessage("Done. Review the preview, then download the updated Excel file.");
    } catch (error) {
      console.error(error);
      setStatus("error");
      setMessage("Something went wrong while processing the Excel file.");
      setProcessedWorkbook(null);
    }
  };

  const handleDownload = () => {
    if (!processedWorkbook) return;
    const outName = fileName.replace(/\.(xlsx|xls|csv)$/i, "") || "processed_addresses";
    XLSX.writeFile(processedWorkbook, `${outName}_formatted.xlsx`);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-slate-100 text-slate-900 p-6 md:p-10">
      <div className="max-w-6xl mx-auto grid gap-6">
        <motion.div
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          className="rounded-3xl border border-slate-200 bg-white/90 shadow-xl p-6 md:p-8"
        >
          <div className="flex items-start justify-between gap-4 flex-col md:flex-row">
            <div>
              <div className="inline-flex items-center gap-2 rounded-full border border-slate-200 px-3 py-1 text-sm font-medium text-slate-600">
                <FileSpreadsheet className="h-4 w-4" />
                Excel Address Splitter
              </div>
              <h1 className="mt-4 text-3xl md:text-4xl font-semibold tracking-tight">
                Upload an Excel file and split Singapore addresses automatically
              </h1>
              <p className="mt-3 text-slate-600 max-w-3xl">
                This app reads one full address column and adds three new columns: Parsed Address, Unit Number, and Postal Code.
              </p>
            </div>
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
                <label className="text-sm font-medium text-slate-700">Column selection mode</label>
                <div className="mt-2 flex gap-2 flex-wrap">
                  <button
                    onClick={() => setColumnMode("header")}
                    className={`rounded-2xl px-4 py-2 border transition ${columnMode === "header" ? "border-slate-900 bg-slate-900 text-white" : "border-slate-300 bg-white text-slate-700"}`}
                  >
                    Use header name
                  </button>
                  <button
                    onClick={() => setColumnMode("letter")}
                    className={`rounded-2xl px-4 py-2 border transition ${columnMode === "letter" ? "border-slate-900 bg-slate-900 text-white" : "border-slate-300 bg-white text-slate-700"}`}
                  >
                    Use column letter
                  </button>
                </div>
              </div>

              {columnMode === "header" ? (
                <div>
                  <label className="text-sm font-medium text-slate-700">Address column header</label>
                  <input
                    value={headerName}
                    onChange={(e) => setHeaderName(e.target.value)}
                    className="mt-2 w-full rounded-2xl border border-slate-300 px-4 py-3 outline-none focus:ring-2 focus:ring-slate-300"
                    placeholder="Example: Address"
                  />
                </div>
              ) : (
                <div>
                  <label className="text-sm font-medium text-slate-700">Address column letter</label>
                  <input
                    value={columnLetter}
                    onChange={(e) => setColumnLetter(e.target.value.toUpperCase())}
                    className="mt-2 w-full rounded-2xl border border-slate-300 px-4 py-3 outline-none focus:ring-2 focus:ring-slate-300"
                    placeholder="Example: A"
                  />
                </div>
              )}

              <div>
                <label className="text-sm font-medium text-slate-700">Excel file</label>
                <label className="mt-2 flex cursor-pointer items-center justify-center gap-3 rounded-3xl border-2 border-dashed border-slate-300 bg-slate-50 px-6 py-10 text-center hover:bg-slate-100 transition">
                  <Upload className="h-5 w-5" />
                  <span className="font-medium">Choose .xlsx, .xls or .csv file</span>
                  <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFileUpload} />
                </label>
                {fileName ? <p className="mt-2 text-sm text-slate-500">Selected file: {fileName}</p> : null}
              </div>

              <div className={`rounded-2xl border p-4 ${status === "error" ? "border-red-200 bg-red-50 text-red-700" : status === "ready" ? "border-emerald-200 bg-emerald-50 text-emerald-700" : "border-slate-200 bg-slate-50 text-slate-600"}`}>
                <div className="flex items-center gap-2 font-medium">
                  {status === "error" ? <AlertCircle className="h-4 w-4" /> : <CheckCircle2 className="h-4 w-4" />}
                  Status
                </div>
                <p className="mt-1 text-sm">{message}</p>
              </div>

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
              New columns will be appended to the right of your existing data.
            </p>

            <div className="mt-5 overflow-auto rounded-2xl border border-slate-200">
              <table className="min-w-full text-sm">
                <tbody>
                  {rowsPreview.length ? (
                    rowsPreview.map((row, rowIndex) => (
                      <tr key={rowIndex} className={rowIndex === 0 ? "bg-slate-100 font-semibold" : "border-t border-slate-200"}>
                        {row.map((cell, cellIndex) => (
                          <td key={cellIndex} className="px-3 py-2 whitespace-nowrap">
                            {String(cell ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td className="px-4 py-8 text-slate-500">
                        No preview yet. Upload a file to see the first few processed rows.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            <div className="mt-5 rounded-2xl bg-slate-50 border border-slate-200 p-4">
              <p className="font-medium">Expected parsing example</p>
              <p className="mt-2 text-sm text-slate-600">
                Input: <span className="font-mono">234 LOR 8 TOA PAYOH #01-284 SINGAPORE 310234</span>
              </p>
              <div className="mt-3 grid md:grid-cols-3 gap-3 text-sm">
                <div className="rounded-2xl bg-white border border-slate-200 p-3">
                  <div className="text-slate-500">Parsed Address</div>
                  <div className="mt-1 font-medium">234 LOR 8 TOA PAYOH</div>
                </div>
                <div className="rounded-2xl bg-white border border-slate-200 p-3">
                  <div className="text-slate-500">Unit Number</div>
                  <div className="mt-1 font-medium">#01-284</div>
                </div>
                <div className="rounded-2xl bg-white border border-slate-200 p-3">
                  <div className="text-slate-500">Postal Code</div>
                  <div className="mt-1 font-medium">310234</div>
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
