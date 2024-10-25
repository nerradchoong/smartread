import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';

function ExcelReader() {
  const [sheets, setSheets] = useState([]);
  const [items, setItems] = useState([]);
  const [formulas, setFormulas] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [workbook, setWorkbook] = useState(null); // State to hold the loaded workbook
  const fileInputRef = useRef(null);

  const interpretFormulas = (formulasRaw, headers) => {
    let formulaInterpretations = {};
    let cachedFormulas = {};  // Cache to store interpretations of formulas

    formulasRaw.forEach(formula => {
      const [cell, formulaStr] = formula.split('=');
      const colIndex = XLSX.utils.decode_cell(cell).c;
      const headerName = headers[colIndex];

      if (cachedFormulas[formulaStr]) { // Check if this formula has already been interpreted
        formulaInterpretations[headerName] = cachedFormulas[formulaStr];
      } else {
        const operationMatch = formulaStr.match(/(SUM|AVERAGE)\(([^)]+)\)/);
        if (operationMatch) {
          const operation = operationMatch[1];
          const range = operationMatch[2];

          const rangeParts = range.split(':');
          if (rangeParts.length === 2) {
            const startCol = XLSX.utils.decode_col(rangeParts[0].match(/[A-Z]+/)[0]);
            const endCol = XLSX.utils.decode_col(rangeParts[1].match(/[A-Z]+/)[0]);
            const colsInvolved = headers.slice(startCol, endCol + 1).join(' + ');

            const interpretation = `${operation}(${colsInvolved})`;
            formulaInterpretations[headerName] = interpretation;
            cachedFormulas[formulaStr] = interpretation; // Cache this new interpretation
          }
        }
      }
    });

    setFormulas(Object.entries(formulaInterpretations).map(([key, value]) => `${key} = ${value}`));
  };

  const readExcel = (file) => {
    if (file && file instanceof Blob) {
      const fileReader = new FileReader();
      fileReader.readAsArrayBuffer(file);

      fileReader.onload = (e) => {
        const bufferArray = e.target.result;
        const wb = XLSX.read(bufferArray, { type: 'buffer', cellFormula: true });
        setWorkbook(wb); // Store the workbook in state
        const sheetNames = wb.SheetNames;
        setSheets(sheetNames);

        if (sheetNames.length > 0) {
          setSelectedSheet(sheetNames[0]);
          extractData(wb, sheetNames[0]);
        }
      };

      fileReader.onerror = (error) => {
        console.error("File could not be read!", error);
      };
    } else {
      console.error("No file selected or the file is not a Blob.");
    }
  };

  const extractData = (workbook, sheetName) => {
    const ws = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
    const formulaData = XLSX.utils.sheet_to_formulae(ws);
    setItems(data);
    interpretFormulas(formulaData, data[0]);
  };

  const handleSheetChange = (e) => {
    const newSheet = e.target.value;
    setSelectedSheet(newSheet);
    if (workbook) {
      extractData(workbook, newSheet); // Use the stored workbook for data extraction
    } else {
      console.error("Workbook is not loaded.");
    }
  };

  return (
    <div>
      <input
        type="file"
        ref={fileInputRef}
        onChange={(e) => {
          const file = e.target.files[0];
          if (file) {
            readExcel(file);
          } else {
            console.error("No file selected.");
          }
        }}
      />
      {sheets.length > 0 && (
        <select onChange={handleSheetChange} value={selectedSheet}>
          {sheets.map((sheet, index) => (
            <option key={index} value={sheet}>
              {sheet}
            </option>
          ))}
        </select>
      )}
      <div className="table-container">
        <table className="table">
          <tbody>
            {items.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, cellIndex) => (
                  <td key={cellIndex}>{cell}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
        {formulas.length > 0 && (
          <>
            <h3>Formulas:</h3>
            <ul>
              {formulas.map((formula, index) => (
                <li key={index}>{formula}</li>
              ))}
            </ul>
          </>
        )}
      </div>
    </div>
  );
}

export default ExcelReader;
