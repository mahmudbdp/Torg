import React, { useState, useEffect, useRef } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { ArrowLeft, UploadCloud, FileSpreadsheet, Plus, RotateCcw, Download, Loader2 } from 'lucide-react';

const OUTPUT_HEADERS = ["DIVISION","SEASON","CUSTOMER","STORE","PO_NUM","CONF_TYPE","START_DATE","END_DATE",
  "STYLE","COLOR_CODE","LBL_CODE","DIMENSION","SIZE_DESC","TOTAL_QTY","DEPARTMENT","ORG_PRICE","ORD_TYPE",
  "UPC","CUST_NAME","ADDRESS1","ADDRESS2","CITY","STATE","ZIPCODE","COUNTRY","ADDR_TYPE","EMAIL","PHONE",
  "ADDRESS3","ADDRESS4","SHIPPER","FRGT_AMT","MISC_AMT","TAX_AMT","TAX_EMPT","OVER_FRGT","BILL_NUM"];

const REQUIRED_INPUT_COLS = ["Market - Store Name","Amount - Order Shipping","Order - Number","Date - Order Date",
  "Item - SKU","Item - Qty","Item - Price","Amount - Paid by Customer","Ship To - Name","Ship To - Address 1",
  "Ship To - Address 2","Ship To - City","Ship To - State","Ship To - Postal Code","Customer Email",
  "Amount - Order Tax","Ship To - Country","Custom - Field 2"];

const DEFAULT_MAPPING = [
  { store: "Zentail Amazon", code: "TOAMA" },
  { store: "Zentail Outdoor Products", code: "TOECO" },
  { store: "New eBay Store", code: "TOEBA" },
  { store: "Zentail Okeechobee Fats", code: "TOECO" },
  { store: "Zentail Samurai Tactical", code: "TOECO" },
  { store: "Zentail Fieldline", code: "TOECO" },
  { store: "Zentail Walmart", code: "TOWAL" },
  { store: "Zentail Timber Hawk", code: "TOECO" },
  { store: "Zentail Mantis Yoga", code: "TOECO" }
];

const MAPPING_KEY = "ss2bc_store_mapping_v1";
const REPLACEMENT_SOURCE_CODE = "TOECO";
const REPLACEMENT_CUSTOMER_CODE = "RREPL";
const REPLACEMENT_THRESHOLD = 3;

// --- Helpers ---
const num = (v) => {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/[^0-9.\-]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};
const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const parseDatePart = (v) => {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "number") {
    const dc = XLSX.SSF.parse_date_code(v);
    if (dc) return new Date(dc.y, dc.m - 1, dc.d);
  }
  const s = String(v || "").trim();
  const datePart = s.split(" ")[0];
  const parts = datePart.split("/");
  if (parts.length === 3) {
    const mo = parseInt(parts[0], 10), da = parseInt(parts[1], 10), yr = parseInt(parts[2], 10);
    if (!isNaN(mo) && !isNaN(da) && !isNaN(yr)) return new Date(yr, mo - 1, da);
  }
  const parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
};
const fmtDate = (d) => {
  if (!d) return "";
  return (d.getMonth() + 1) + "/" + d.getDate() + "/" + d.getFullYear();
};

const SKU_OVERRIDES = {
  "STT002008ECLRG": { style: "STT002", lbl: "EC", color: "008" },
  "STT002008ECMED": { style: "STT002", lbl: "EC", color: "008" },
  "STT001008ECXLG": { style: "STT001", lbl: "EC", color: "008" },
  "STT001E250ECXLG": { style: "STT001", lbl: "EC", color: "E250" }
};

const splitSku = (rawSku) => {
  let sku = (rawSku || "").trim();
  if (!sku) return { style: "", lbl: "-", color: "", ok: false };
  const fbaMatch = sku.match(/^(.*)-\s*FBA$/i);
  if (fbaMatch) sku = fbaMatch[1].trim();

  const override = SKU_OVERRIDES[sku.toUpperCase()];
  if (override) return { style: override.style, lbl: override.lbl, color: override.color, ok: true };

  const segments = sku.split("-").map(s => s.trim());
  if (segments.length === 1) return { style: sku, lbl: "-", color: "", ok: false };

  let prefix, color;
  if (segments.length >= 3 && segments[1] === "") {
    prefix = segments[0];
    color = segments[2];
    return { style: prefix, lbl: "-", color: color, ok: true };
  }

  prefix = segments[0];
  color = segments[1];

  let style, lbl, ok = true;
  const m = prefix.match(/^(.*\d)([A-Za-z]*)$/);
  if (m) {
    style = m[1];
    lbl = m[2] || "-";
  } else if (prefix.length > 2) {
    style = prefix.slice(0, -2);
    lbl = prefix.slice(-2);
  } else {
    style = prefix;
    lbl = "-";
    ok = false;
  }
  return { style, lbl, color, ok };
};

const buildHeaderIndex = (headerRow) => {
  const idx = {};
  headerRow.forEach((h, i) => { idx[String(h || "").trim().toLowerCase()] = i; });
  return idx;
};
const getCell = (row, idx, name) => {
  const i = idx[name.toLowerCase()];
  if (i === undefined) return "";
  const v = row[i];
  return v === undefined || v === null ? "" : v;
};

// --- Component ---
const Converter = () => {
  useEffect(() => {
    document.title = "Converter | TORG ODP JOB";
  }, []);

  const [currentStep, setCurrentStep] = useState(1);
  const [mapping, setMapping] = useState(DEFAULT_MAPPING);
  const [loadedRows, setLoadedRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [isDrag, setIsDrag] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  
  const [messages, setMessages] = useState([]);
  const [lastOutput, setLastOutput] = useState(null);
  
  const fileInputRef = useRef(null);

  useEffect(() => {
    try {
      const saved = localStorage.getItem(MAPPING_KEY);
      if (saved) setMapping(JSON.parse(saved));
    } catch (e) {}
  }, []);

  const saveMapping = (newMapping) => {
    setMapping(newMapping);
    try { localStorage.setItem(MAPPING_KEY, JSON.stringify(newMapping)); } catch(e){}
  };

  const handleMappingChange = (i, field, value) => {
    const updated = [...mapping];
    updated[i][field] = value;
    saveMapping(updated);
  };
  const addMappingRow = () => saveMapping([...mapping, { store: "", code: "" }]);
  const removeMappingRow = (i) => {
    const updated = [...mapping];
    updated.splice(i, 1);
    saveMapping(updated);
  };
  const resetMapping = () => {
    if (window.confirm("Reset store mapping to defaults?")) saveMapping(DEFAULT_MAPPING);
  };

  const processFile = (file) => {
    setFileName(file.name);
    setMessages([]);
    setLastOutput(null);
    
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        let rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        rows = rows.filter(r => r.some(c => String(c).trim() !== ""));
        if (!rows.length) { 
          setMessages([{ type: 'err', text: 'File appears empty.' }]);
          return;
        }
        setLoadedRows(rows);
        setMessages([{ type: 'ok', text: `Loaded ${rows.length - 1} data row(s) from ${file.name}.` }]);
      } catch (err) {
        setMessages([{ type: 'err', text: `Could not read file: ${err.message}` }]);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const onConvert = () => {
    if (!loadedRows) return;
    setIsConverting(true);
    setMessages([]);
    setLastOutput(null);

    // Artificial delay for "App" feel
    setTimeout(() => {
      runConversionLogic();
    }, 800);
  };

  const runConversionLogic = () => {
    const header = loadedRows[0];
    const idx = buildHeaderIndex(header);
    const msgs = [];

    const missing = REQUIRED_INPUT_COLS.filter(c => idx[c.toLowerCase()] === undefined);
    if (missing.length) {
      msgs.push({ type: 'err', text: `Missing expected ShipStation column(s): ${missing.join(", ")}` });
      setMessages(msgs);
      setIsConverting(false);
      return;
    }

    const dataRows = loadedRows.slice(1);
    const todayStr = fmtDate(new Date());
    
    const mapDict = {};
    mapping.forEach(e => { if (e.store) mapDict[e.store.trim().toLowerCase()] = e.code || ""; });

    const qtyByOrder = {};
    dataRows.forEach(r => {
      const ordNum = String(getCell(r, idx, "Order - Number")).trim();
      qtyByOrder[ordNum] = (qtyByOrder[ordNum] || 0) + num(getCell(r, idx, "Item - Qty"));
    });

    const unmappedStores = new Set();
    const badSkus = [];
    const outRows = [];
    const orderLevelAssigned = new Set();

    dataRows.forEach((r, rowIdx) => {
      const ordNum = String(getCell(r, idx, "Order - Number")).trim();
      if (!ordNum) return;

      const store = String(getCell(r, idx, "Market - Store Name")).trim();
      let customerCode = mapDict[store.toLowerCase()];
      if (customerCode === undefined) {
        unmappedStores.add(store);
        customerCode = "";
      }

      const sku = getCell(r, idx, "Item - SKU");
      const skuParts = splitSku(sku);
      if (!skuParts.ok) badSkus.push(`Row ${rowIdx + 2}: SKU "${sku}" (order ${ordNum})`);

      const qty = num(getCell(r, idx, "Item - Qty"));
      const paid = num(getCell(r, idx, "Amount - Paid by Customer"));
      const tax = num(getCell(r, idx, "Amount - Order Tax"));
      const freight = num(getCell(r, idx, "Amount - Order Shipping"));
      const totalQty = qtyByOrder[ordNum] || qty || 1;
      
      const isToecoOrder = customerCode === REPLACEMENT_SOURCE_CODE;
      const orgPrice = isToecoOrder
        ? round2((paid - tax - freight) / totalQty)
        : round2((paid - tax) / totalQty);

      if (isToecoOrder && round2(paid - freight) < REPLACEMENT_THRESHOLD) {
        customerCode = REPLACEMENT_CUSTOMER_CODE;
      }

      const custField2 = String(getCell(r, idx, "Custom - Field 2")).trim();
      const poNum = custField2 || ordNum;
      const custName = String(getCell(r, idx, "Ship To - Name")).replace(/&/g, "and");
      const orderDate = parseDatePart(getCell(r, idx, "Date - Order Date"));

      const out = {};
      out.DIVISION = 1;
      out.SEASON = "ALL";
      out.CUSTOMER = customerCode;
      out.STORE = "";
      out.PO_NUM = poNum;
      out.CONF_TYPE = "A";
      out.START_DATE = fmtDate(orderDate);
      out.END_DATE = todayStr;
      out.STYLE = skuParts.style;
      out.COLOR_CODE = skuParts.color;
      out.LBL_CODE = skuParts.lbl;
      out.DIMENSION = "-";
      out.SIZE_DESC = "-";
      out.TOTAL_QTY = qty;
      out.DEPARTMENT = "";
      out.ORG_PRICE = orgPrice;
      out.ORD_TYPE = "REG";
      out.UPC = "";
      out.CUST_NAME = custName;
      out.ADDRESS1 = getCell(r, idx, "Ship To - Address 1");
      out.ADDRESS2 = getCell(r, idx, "Ship To - Address 2");
      out.CITY = getCell(r, idx, "Ship To - City");
      out.STATE = getCell(r, idx, "Ship To - State");
      out.ZIPCODE = getCell(r, idx, "Ship To - Postal Code");
      out.COUNTRY = getCell(r, idx, "Ship To - Country");
      out.ADDR_TYPE = "OT";
      out.EMAIL = getCell(r, idx, "Customer Email");
      out.PHONE = "";
      out.ADDRESS3 = "";
      out.ADDRESS4 = "";
      out.SHIPPER = "F11";
      if (orderLevelAssigned.has(ordNum)) {
        out.FRGT_AMT = 0;
        out.TAX_AMT = 0;
      } else {
        out.FRGT_AMT = freight;
        out.TAX_AMT = tax;
        orderLevelAssigned.add(ordNum);
      }
      out.MISC_AMT = 0;
      out.TAX_EMPT = "P";
      out.OVER_FRGT = "P";
      out.BILL_NUM = ordNum;

      outRows.push(OUTPUT_HEADERS.map(h => out[h]));
    });

    if (!outRows.length) {
      msgs.push({ type: 'err', text: "No valid data rows found to convert." });
      setMessages(msgs);
      setIsConverting(false);
      return;
    }

    const aoa = [OUTPUT_HEADERS, ...outRows];
    setLastOutput(aoa);

    msgs.push({ type: 'ok', text: `Converted ${outRows.length} row(s). Review the preview below, then download.` });
    
    if (unmappedStores.size) {
      msgs.push({ type: 'warn', title: "Missing CUSTOMER mapping — add them in step 2:", items: Array.from(unmappedStores) });
    }
    if (badSkus.length) {
      msgs.push({ type: 'warn', title: "SKUs didn't match expected pattern (STYLE+LBLCODE-COLOR):", items: badSkus });
    }

    setMessages(msgs);
    setIsConverting(false);
  };

  const onDownload = () => {
    if (!lastOutput) return;
    const ws = XLSX.utils.aoa_to_sheet(lastOutput);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MJSO");
    const now = new Date();
    const stamp = String(now.getMonth()+1).padStart(2,"0") + String(now.getDate()).padStart(2,"0") + 
                  now.getFullYear() + String(now.getHours()).padStart(2,"0") + String(now.getMinutes()).padStart(2,"0");
    XLSX.writeFile(wb, `MJSO_EXCEL_${stamp}.xls`, { bookType: "xls" });
  };

  return (
    <div className="animate-fade-in bg-gray-50 min-h-screen pb-20">
      
      {/* Hero Header */}
      <div className="bg-gradient-to-br from-white via-blue-50 to-orange-50/30 text-aws-navy pt-10 pb-20 px-6 relative overflow-hidden border-b border-gray-200">
        <div className="absolute top-0 right-0 w-96 h-96 bg-blue-500/10 blur-[100px] rounded-full pointer-events-none" />
        <div className="max-w-4xl mx-auto relative z-10">
          <Link to="/" className="inline-flex items-center text-sm font-medium text-aws-blue hover:text-blue-800 transition-colors mb-6">
            <ArrowLeft size={16} className="mr-2" /> Back to all tools
          </Link>
          <h1 className="text-3xl font-extrabold mb-3">ShipStation to Bluecherry (MJSO) Converter</h1>
          <p className="text-gray-500 text-lg">Converts order exports locally securely inside your browser.</p>
        </div>
      </div>

      <div className="max-w-4xl mx-auto px-6 -mt-10 relative z-20 space-y-6">
        
        {/* Step Indicator */}
        <div className="glass-card p-6 flex justify-between items-center mb-6">
          {[1, 2, 3].map((step) => (
            <div key={step} className={`flex items-center ${step < 3 ? 'flex-1' : ''}`}>
              <div className={`w-10 h-10 rounded-full flex items-center justify-center font-bold text-lg shadow-md transition-all ${
                currentStep === step 
                  ? 'bg-aws-orange text-white' 
                  : currentStep > step 
                    ? 'bg-green-500 text-white' 
                    : 'bg-gray-200 text-gray-500'
              }`}>
                {currentStep > step ? '✓' : step}
              </div>
              {step < 3 && (
                <div className={`h-1 flex-1 mx-4 rounded-full transition-all ${
                  currentStep > step ? 'bg-green-500' : 'bg-gray-200'
                }`} />
              )}
            </div>
          ))}
        </div>

        <div className="glass-card p-8 min-h-[400px] flex flex-col">
          {/* Step 1 Content */}
          {currentStep === 1 && (
            <div className="flex-grow animate-fade-in">
              <h2 className="text-xl font-bold mb-6 text-aws-navy">Upload ShipStation Order Export</h2>
              <div 
                className={`border-2 border-dashed rounded-xl p-12 text-center cursor-pointer transition-all ${isDrag ? 'border-aws-orange bg-orange-50/50' : 'border-gray-300 hover:border-aws-blue bg-gray-50/50'}`}
                onDragOver={(e) => { e.preventDefault(); setIsDrag(true); }}
                onDragLeave={() => setIsDrag(false)}
                onDrop={(e) => { e.preventDefault(); setIsDrag(false); if(e.dataTransfer.files.length) processFile(e.dataTransfer.files[0]); }}
                onClick={() => fileInputRef.current?.click()}
              >
                <input type="file" ref={fileInputRef} className="hidden" accept=".csv,.xlsx,.xls" onChange={(e) => e.target.files.length && processFile(e.target.files[0])} />
                <UploadCloud size={64} className={`mx-auto mb-6 ${isDrag ? 'text-aws-orange' : 'text-gray-400'}`} />
                <p className="text-gray-600 font-medium text-lg">Click to choose file, or drag & drop here</p>
                <p className="text-gray-400 text-sm mt-2">Supports .csv, .xlsx, .xls</p>
                {fileName && <div className="mt-6 inline-flex items-center bg-white border border-green-200 shadow-sm px-6 py-3 rounded-lg text-md font-bold text-green-700"><FileSpreadsheet size={20} className="mr-2"/>{fileName} successfully loaded!</div>}
              </div>
            </div>
          )}

          {/* Step 2 Content */}
          {currentStep === 2 && (
            <div className="flex-grow animate-fade-in">
              <div className="flex justify-between items-center mb-6">
                 <h2 className="text-xl font-bold text-aws-navy">Store to Customer Code Mapping</h2>
                <span className="text-xs font-semibold bg-blue-50 text-aws-blue px-3 py-1 rounded-full border border-blue-100 shadow-sm">Auto-saves to browser</span>
              </div>
             
              <div className="bg-gray-50 border border-gray-200 rounded-xl overflow-hidden mb-4">
                <div className="max-h-[300px] overflow-y-auto p-4 space-y-3">
                  {mapping.map((row, i) => (
                    <div key={i} className="flex gap-3 items-center group">
                      <input 
                        className="flex-1 px-4 py-3 bg-white border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-aws-blue/50 focus:border-aws-blue transition-all shadow-sm"
                        placeholder="Market - Store Name (exact)"
                        value={row.store}
                        onChange={(e) => handleMappingChange(i, 'store', e.target.value)}
                      />
                      <input 
                        className="w-48 px-4 py-3 bg-white border border-gray-300 rounded-lg text-sm font-mono focus:outline-none focus:ring-2 focus:ring-aws-blue/50 focus:border-aws-blue transition-all shadow-sm"
                        placeholder="CUSTOMER code"
                        value={row.code}
                        onChange={(e) => handleMappingChange(i, 'code', e.target.value)}
                      />
                      <button onClick={() => removeMappingRow(i)} className="text-red-400 hover:text-red-600 p-2 opacity-0 group-hover:opacity-100 transition-opacity">
                         &times;
                      </button>
                    </div>
                  ))}
                </div>
              </div>
              
              <div className="flex gap-3">
                <button onClick={addMappingRow} className="flex items-center text-sm font-semibold text-aws-navy bg-gray-100 hover:bg-gray-200 px-4 py-2 rounded-lg transition-colors border border-gray-300 shadow-sm">
                  <Plus size={16} className="mr-1" /> Add Row
                </button>
                <button onClick={resetMapping} className="flex items-center text-sm font-semibold text-gray-500 hover:text-gray-700 bg-transparent hover:bg-gray-100 px-4 py-2 rounded-lg transition-colors">
                  <RotateCcw size={16} className="mr-1" /> Reset Defaults
                </button>
              </div>
            </div>
          )}

          {/* Step 3 Content */}
          {currentStep === 3 && (
            <div className="flex-grow animate-fade-in">
              <h2 className="text-xl font-bold mb-6 text-aws-navy">Convert & Download</h2>
              
              <div className="flex flex-wrap gap-4 items-center mb-8">
                <button 
                  onClick={onConvert} 
                  disabled={!loadedRows || isConverting}
                  className="bg-aws-orange hover:bg-orange-600 disabled:bg-gray-300 disabled:cursor-not-allowed text-white font-bold py-3 px-8 rounded-lg transition-all shadow-md flex items-center gap-2 text-lg"
                >
                  {isConverting ? <Loader2 className="animate-spin" size={24} /> : <FileSpreadsheet size={24} />}
                  Convert File Now
                </button>
                
                {lastOutput && (
                  <button 
                    onClick={onDownload}
                    className="bg-aws-blue hover:bg-blue-700 text-white font-bold py-3 px-8 rounded-lg transition-all shadow-md flex items-center gap-2 text-lg animate-fade-in"
                  >
                    <Download size={24} /> Download MJSO .xls
                  </button>
                )}
              </div>

              {/* Messages */}
              {messages.length > 0 && (
                <div className="space-y-3 mb-6 animate-fade-in">
                  {messages.map((m, i) => (
                    <div key={i} className={`p-4 rounded-xl border-l-4 shadow-sm text-sm ${
                      m.type === 'err' ? 'bg-red-50 border-red-500 text-red-900' :
                      m.type === 'ok' ? 'bg-green-50 border-green-500 text-green-900' :
                      'bg-orange-50 border-aws-orange text-orange-900'
                    }`}>
                      {m.text && <div className="font-semibold">{m.text}</div>}
                      {m.title && <div className="font-semibold mb-2">{m.title}</div>}
                      {m.items && (
                        <ul className="list-disc pl-5 space-y-1 mt-2 text-xs opacity-80 max-h-40 overflow-y-auto">
                          {m.items.slice(0,25).map((it, j) => <li key={j}>{it}</li>)}
                          {m.items.length > 25 && <li>...and {m.items.length - 25} more</li>}
                        </ul>
                      )}
                    </div>
                  ))}
                </div>
              )}

              {/* Preview Table */}
              {lastOutput && (
                <div className="border border-gray-200 rounded-xl overflow-hidden shadow-sm bg-white animate-fade-in">
                  <div className="bg-gray-50 px-4 py-3 border-b border-gray-200 text-xs font-bold text-gray-500 uppercase tracking-wider flex justify-between">
                    <span>Preview (First 50 Rows)</span>
                    <span>Total: {lastOutput.length - 1} rows</span>
                  </div>
                  <div className="overflow-x-auto max-h-[300px]">
                    <table className="w-full text-left text-xs whitespace-nowrap">
                      <thead className="bg-gray-100 sticky top-0 shadow-sm z-10">
                        <tr>
                          {lastOutput[0].map((h, i) => <th key={i} className="px-3 py-2 font-bold text-gray-700 border-b border-gray-200">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {lastOutput.slice(1, 51).map((row, i) => (
                          <tr key={i} className="hover:bg-gray-50 border-b border-gray-100 last:border-0">
                            {row.map((cell, j) => <td key={j} className="px-3 py-2 text-gray-600">{cell}</td>)}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Navigation Buttons */}
          <div className="mt-8 pt-6 border-t border-gray-200 flex justify-between items-center">
            {currentStep === 1 ? (
              <Link
                to="/"
                className="font-semibold py-2 px-6 rounded-lg transition-all bg-gray-100 hover:bg-gray-200 text-gray-700 border border-gray-300 shadow-sm inline-flex items-center"
              >
                <ArrowLeft size={16} className="mr-2" /> Back to Home
              </Link>
            ) : (
              <button
                onClick={() => setCurrentStep(prev => Math.max(1, prev - 1))}
                className="font-semibold py-2 px-6 rounded-lg transition-all bg-gray-100 hover:bg-gray-200 text-gray-700 border border-gray-300 shadow-sm inline-flex items-center"
              >
                <ArrowLeft size={16} className="mr-2" /> Back
              </button>
            )}

            <div className="flex gap-4">
              <button
                onClick={() => {
                  setCurrentStep(1);
                  setLoadedRows(null);
                  setFileName("");
                  setMessages([]);
                  setLastOutput(null);
                  if (fileInputRef.current) fileInputRef.current.value = "";
                }}
                className={`font-semibold py-2 px-6 rounded-lg transition-all bg-white hover:bg-gray-50 text-gray-600 border border-gray-300 shadow-sm inline-flex items-center ${!loadedRows ? 'opacity-0 pointer-events-none' : ''}`}
              >
                <RotateCcw size={16} className="mr-2" /> Start Over
              </button>

              {currentStep < 3 && (
                <button
                  onClick={() => setCurrentStep(prev => Math.min(3, prev + 1))}
                  disabled={currentStep === 1 && !loadedRows}
                  className={`font-semibold py-2 px-6 rounded-lg transition-all ${(currentStep === 1 && !loadedRows) ? 'bg-gray-300 text-gray-500 cursor-not-allowed' : 'bg-aws-navy hover:bg-aws-squid-ink text-white shadow-md'}`}
                >
                  Next Step
                </button>
              )}
            </div>
          </div>

        </div>

      </div>
    </div>
  );
};

export default Converter;
