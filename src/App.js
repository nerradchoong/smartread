import React from 'react';
import './App.css';
import ExcelReader from './ExcelReader';
//import ExcelReader from './ExcelReaderBackup';
import axios from 'axios';

// Set withCredentials globally
axios.defaults.withCredentials = true;

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <p>
          Upload and Read Excel File
        </p>
        <ExcelReader />
      </header>
    </div>
  );
}

export default App;
