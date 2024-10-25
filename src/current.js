import React, { useState, useRef } from 'react';
import axios from 'axios';
import * as XLSX from 'xlsx';

function ExcelReader() {
  const [sheets, setSheets] = useState([]);
  const [items, setItems] = useState([]);
  const [formulas, setFormulas] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [workbook, setWorkbook] = useState(null); // State to hold the loaded workbook
  const fileInputRef = useRef(null);

  const fetchFormulas = async (worksheet, headers) => {
    const formulaData = XLSX.utils.sheet_to_formulae(worksheet);
    try {
      const response = await axios.post('http://localhost:3001/api/excel/interpretFormulas', {
        formulas: formulaData,
        headers: headers
      }, { withCredentials: true });
      if (response.data) {
        setFormulas(response.data.formulas || []);
      }
    } catch (error) {
      console.error('Error fetching formulas:', error.response ? error.response.data : error.message);
    }
  };

  const readExcel = (file) => {
    if (file && file instanceof Blob) {
      const fileReader = new FileReader();
      fileReader.readAsArrayBuffer(file);

      fileReader.onload = async (e) => {
        const bufferArray = e.target.result;
        const wb = XLSX.read(bufferArray, { type: 'buffer', cellFormula: true });
        setWorkbook(wb); // Store the workbook in state
        const sheetNames = wb.SheetNames;
        setSheets(sheetNames);

        if (sheetNames.length > 0) {
          setSelectedSheet(sheetNames[0]);
          const ws = wb.Sheets[sheetNames[0]];
          const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
          setItems(data);
          await fetchFormulas(ws, data[0]);
        }
      };

      fileReader.onerror = (error) => {
        console.error("File could not be read!", error);
      };
    } else {
      console.error("No file selected or the file is not a Blob.");
    }
  };

  const handleSheetChange = async (e) => {
    const newSheet = e.target.value;
    setSelectedSheet(newSheet);
    if (workbook) {
      const ws = workbook.Sheets[newSheet];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
      setItems(data);
      await fetchFormulas(ws, data[0]); // Fetch formulas for the new sheet
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
                <li key={index}>{formula.header} = {formula.formula}</li>
              ))}
            </ul>
          </>
        )}
      </div>
    </div>
  );
}

export default ExcelReader;
