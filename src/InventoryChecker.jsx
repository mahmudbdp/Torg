import React, { useState, useRef } from 'react';
import { Link } from 'react-router-dom';
import { ArrowLeft, UploadCloud, FileSpreadsheet, CheckCircle, AlertCircle, Loader2 } from 'lucide-react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const InventoryChecker = () => {
  const [salesFile, setSalesFile] = useState(null);
  const [inventoryFile, setInventoryFile] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [previewData, setPreviewData] = useState(null);
  const [excelBlob, setExcelBlob] = useState(null);

  const handleSalesUpload = (e) => {
    if (e.target.files && e.target.files[0]) {
      setSalesFile(e.target.files[0]);
      setError('');
    }
  };

  const handleInventoryUpload = (e) => {
    if (e.target.files && e.target.files[0]) {
      setInventoryFile(e.target.files[0]);
      setError('');
    }
  };

  const processFiles = async () => {
    if (!salesFile || !inventoryFile) {
      setError('Please upload both files before processing.');
      return;
    }

    setIsProcessing(true);
    setMessage('Reading Inventory file...');
    setError('');

    try {
      // 1. Read Inventory File
      const invData = await readFileAsync(inventoryFile);
      const invWorkbook = XLSX.read(invData, { type: 'array' });
      const invSheet = invWorkbook.Sheets[invWorkbook.SheetNames[0]];
      const invRows = XLSX.utils.sheet_to_json(invSheet, { header: 1 });

      if (invRows.length < 2) throw new Error("Inventory file is empty.");

      const invHeaders = invRows[0];
      const qtyIdx = invHeaders.findIndex(h => h && h.toString().trim() === 'Total Avail QOH');
      const eskuIdx = invHeaders.findIndex(h => h && h.toString().trim() === 'ESKU');
      const locIdx = invHeaders.findIndex(h => h && h.toString().trim() === 'Loc');

      if (qtyIdx === -1 || eskuIdx === -1 || locIdx === -1) {
        throw new Error("Inventory file missing required columns: 'Total Avail QOH', 'ESKU', or 'Loc'.");
      }

      // Build dictionary: dict[ESKU][LOC] = sum(Qty)
      const inventoryDict = {};
      for (let i = 1; i < invRows.length; i++) {
        const row = invRows[i];
        if (!row || row.length === 0) continue;
        const esku = row[eskuIdx];
        const loc = row[locIdx];
        const qty = parseFloat(row[qtyIdx]) || 0;

        if (esku && loc) {
          const locStr = loc.toString().toUpperCase();
          let baseLoc = null;
          if (locStr.includes('JMRA')) baseLoc = 'JMRA';
          else if (locStr.includes('WTHA')) baseLoc = 'WTHA';
          else if (locStr.includes('WTHS')) baseLoc = 'WTHS';
          else if (locStr.includes('WTHW')) baseLoc = 'WTHW';

          if (baseLoc) {
            if (!inventoryDict[esku]) inventoryDict[esku] = {};
            inventoryDict[esku][baseLoc] = (inventoryDict[esku][baseLoc] || 0) + qty;
          }
        }
      }

      setMessage('Reading Sales Order file...');
      
      // 2. Read Sales Order File
      const salesData = await readFileAsync(salesFile);
      const salesWorkbook = XLSX.read(salesData, { type: 'array', cellDates: true });
      const salesSheet = salesWorkbook.Sheets[salesWorkbook.SheetNames[0]];
      const salesRows = XLSX.utils.sheet_to_json(salesSheet, { header: 1, raw: false, dateNF: 'mm/dd/yyyy hh:mm:ss AM/PM' });

      if (salesRows.length < 2) throw new Error("Sales file is empty.");

      const salesHeaders = salesRows[0];
      const styleIdx = salesHeaders.findIndex(h => h && h.toString().trim() === 'Style');
      const labelIdx = salesHeaders.findIndex(h => h && h.toString().trim() === 'Label');
      const colorIdx = salesHeaders.findIndex(h => h && h.toString().trim() === 'Color Code');

      if (styleIdx === -1 || labelIdx === -1 || colorIdx === -1) {
        throw new Error("Sales file missing required columns: 'Style', 'Label', or 'Color Code'.");
      }

      setMessage('Generating Output Excel...');

      // 3. Create ExcelJS Workbook for styling
      const outWorkbook = new ExcelJS.Workbook();
      const outSheet = outWorkbook.addWorksheet('Sales Order Detail');

      // Setup Headers
      const newHeaders = [...salesHeaders];
      // Insert ESKU exactly after Color Code
      const insertIdx = colorIdx + 1;
      newHeaders.splice(insertIdx, 0, 'ESKU');
      
      // Add Location columns at the end
      newHeaders.push('JMRA', 'WTHA', 'WTHS', 'WTHW');
      
      outSheet.addRow(newHeaders);

      // Make headers bold
      outSheet.getRow(1).font = { bold: true };

      // Process Rows
      for (let i = 1; i < salesRows.length; i++) {
        const row = salesRows[i];
        if (!row || row.length === 0) continue;

        const style = row[styleIdx] || '';
        const label = row[labelIdx] || '';
        const color = row[colorIdx] || '';

        // Generate ESKU
        let esku = '';
        if (style || label || color) {
          esku = `${style}${label}-${color}`;
        }

        const newRow = [...row];
        // Insert ESKU
        newRow.splice(insertIdx, 0, esku);

        // Fetch Inventory
        const invEntry = inventoryDict[esku] || {};
        const jmra = invEntry['JMRA'] !== undefined ? invEntry['JMRA'] : '#N/A';
        const wtha = invEntry['WTHA'] !== undefined ? invEntry['WTHA'] : '#N/A';
        const wths = invEntry['WTHS'] !== undefined ? invEntry['WTHS'] : '#N/A';
        const wthw = invEntry['WTHW'] !== undefined ? invEntry['WTHW'] : '#N/A';

        // Add to row
        newRow.push(jmra, wtha, wths, wthw);

        const excelRow = outSheet.addRow(newRow);

        // Apply Styles to the 4 Location Columns (which are the last 4 columns)
        const colCount = newRow.length;
        for (let c = colCount - 3; c <= colCount; c++) {
          const cell = excelRow.getCell(c);
          const val = cell.value;
          
          if (val !== '#N/A') {
            if (parseFloat(val) >= 1) {
              // Green: Light green background, dark green text
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD4EDDA' } };
              cell.font = { color: { argb: 'FF155724' } };
            } else {
              // Red: Light red background, dark red text
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8D7DA' } };
              cell.font = { color: { argb: 'FF721C24' } };
            }
          }
        }
      }

      // Build preview data (first 50 rows max to keep it fast)
      const previewRows = [];
      outSheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 50) {
          previewRows.push(row.values.slice(1)); // .values is 1-indexed array in exceljs
        }
      });
      setPreviewData({ headers: newHeaders, rows: previewRows.slice(1) }); // exclude header row from rows

      // Save Blob
      const buffer = await outWorkbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      setExcelBlob(blob);

      setMessage('Process Complete! Review the data below and download the file.');
      setSalesFile(null);
      setInventoryFile(null);
    } catch (err) {
      console.error(err);
      setError(err.message || 'An error occurred during processing.');
    } finally {
      setIsProcessing(false);
    }
  };

  const readFileAsync = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(new Uint8Array(e.target.result));
      reader.onerror = (e) => reject(e);
      reader.readAsArrayBuffer(file);
    });
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
          <h1 className="text-3xl font-extrabold mb-3">Inventory Checker</h1>
          <p className="text-gray-500 text-lg">Merge Open Pick Sales Orders with Inventory Quantities.</p>
        </div>
      </div>

      <div className="max-w-4xl mx-auto px-6 -mt-10 relative z-20">
        <div className="bg-white rounded-2xl shadow-xl border border-gray-100 p-8">
          
          {error && (
            <div className="mb-8 p-4 bg-red-50 border-l-4 border-red-500 text-red-700 rounded-md flex items-start gap-3">
              <AlertCircle className="flex-shrink-0 mt-0.5" size={20} />
              <div>
                <h3 className="font-bold text-sm">Processing Error</h3>
                <p className="text-sm">{error}</p>
              </div>
            </div>
          )}

          {message && !error && (
            <div className="mb-8 p-4 bg-green-50 border border-green-200 text-green-700 rounded-md flex items-center gap-3">
              <CheckCircle size={20} />
              <p className="font-medium">{message}</p>
            </div>
          )}

          <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-8">
            {/* Sales File Upload */}
            <div className="border-2 border-dashed border-gray-300 rounded-xl p-6 text-center hover:bg-gray-50 transition-colors relative">
              <input 
                type="file" 
                accept=".xlsx,.xls,.csv"
                onChange={handleSalesUpload}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
              />
              <div className="flex flex-col items-center justify-center h-full pointer-events-none">
                <div className={`w-12 h-12 rounded-full flex items-center justify-center mb-3 ${salesFile ? 'bg-green-100 text-green-600' : 'bg-blue-100 text-blue-600'}`}>
                  {salesFile ? <CheckCircle size={24} /> : <FileSpreadsheet size={24} />}
                </div>
                <h3 className="font-bold text-gray-900 mb-1">1. Sales Order Detail</h3>
                <p className="text-xs text-gray-500 px-4">
                  {salesFile ? salesFile.name : 'Drag & drop or click to upload the Open Pick report'}
                </p>
              </div>
            </div>

            {/* Inventory File Upload */}
            <div className="border-2 border-dashed border-gray-300 rounded-xl p-6 text-center hover:bg-gray-50 transition-colors relative">
              <input 
                type="file" 
                accept=".xlsx,.xls,.csv"
                onChange={handleInventoryUpload}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
              />
              <div className="flex flex-col items-center justify-center h-full pointer-events-none">
                <div className={`w-12 h-12 rounded-full flex items-center justify-center mb-3 ${inventoryFile ? 'bg-green-100 text-green-600' : 'bg-orange-100 text-aws-orange'}`}>
                  {inventoryFile ? <CheckCircle size={24} /> : <UploadCloud size={24} />}
                </div>
                <h3 className="font-bold text-gray-900 mb-1">2. Inventory Report</h3>
                <p className="text-xs text-gray-500 px-4">
                  {inventoryFile ? inventoryFile.name : 'Drag & drop or click to upload the Inventory report'}
                </p>
              </div>
            </div>
          </div>

          <div className="flex justify-end">
            <button
              onClick={processFiles}
              disabled={isProcessing || !salesFile || !inventoryFile}
              className={`py-3 px-8 rounded-lg font-bold flex items-center gap-2 transition-all shadow-md ${
                isProcessing || !salesFile || !inventoryFile
                  ? 'bg-gray-300 text-gray-500 cursor-not-allowed'
                  : 'bg-aws-orange text-white hover:bg-orange-600 hover:shadow-lg'
              }`}
            >
              {isProcessing ? (
                <>
                  <Loader2 size={18} className="animate-spin" /> Processing...
                </>
              ) : (
                'Generate Preview'
              )}
            </button>
          </div>

          {/* Preview Section */}
          {previewData && excelBlob && (
            <div className="mt-12 border-t border-gray-200 pt-8 animate-fade-in">
              <div className="flex items-center justify-between mb-6">
                <div>
                  <h2 className="text-2xl font-bold text-gray-900 mb-1">Data Preview</h2>
                  <p className="text-gray-500 text-sm">Showing the first {previewData.rows.length} rows of the generated output.</p>
                </div>
                <button
                  onClick={() => saveAs(excelBlob, `Sales Order Detail - Open Pick_report_${new Date().getTime()}.xlsx`)}
                  className="bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-6 rounded-lg shadow-md transition-all flex items-center gap-2"
                >
                  <FileSpreadsheet size={20} />
                  Download Excel File
                </button>
              </div>

              <div className="overflow-x-auto border border-gray-200 rounded-lg shadow-sm">
                <table className="w-full text-sm text-left">
                  <thead className="bg-gray-50 text-gray-700 uppercase font-semibold text-xs sticky top-0">
                    <tr>
                      {previewData.headers.map((h, i) => (
                        <th key={i} className="px-4 py-3 whitespace-nowrap border-b border-gray-200">{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {previewData.rows.map((row, i) => (
                      <tr key={i} className="bg-white border-b hover:bg-gray-50">
                        {previewData.headers.map((_, colIdx) => {
                          const val = row[colIdx];
                          let cellClass = "px-4 py-3 whitespace-nowrap";
                          
                          // Style location columns (last 4 cols)
                          if (colIdx >= previewData.headers.length - 4) {
                            if (val !== undefined && val !== '#N/A') {
                              if (parseFloat(val) >= 1) {
                                cellClass += " bg-green-100 text-green-800 font-medium";
                              } else {
                                cellClass += " bg-red-100 text-red-800 font-medium";
                              }
                            }
                          }
                          
                          return (
                            <td key={colIdx} className={cellClass}>
                              {val !== undefined ? val.toString() : ''}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

        </div>
      </div>
    </div>
  );
};

export default InventoryChecker;
