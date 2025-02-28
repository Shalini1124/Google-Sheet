import React from "react";
import Spreadsheet from "./components/spreadsheet";
import "./App.css"; 

function App() {
  return (
    <div className="app-container">
      <h1 className="app-title">Google Sheets Clone</h1>
      <Spreadsheet />
    </div>
  );
}

export default App;
