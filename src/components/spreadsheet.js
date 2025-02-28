import React, { useState, useRef, useEffect } from "react";
import Cell from "./cell";
import "./spreadsheet.css";

const DEFAULT_ROWS = 20;
const DEFAULT_COLS = 10;

const Spreadsheet = () => {
  // Initialize state for rows and columns
  const [rowCount, setRowCount] = useState(DEFAULT_ROWS);
  const [colCount, setColCount] = useState(DEFAULT_COLS);
  
  // Initialize cells with empty values
  const [cells, setCells] = useState(
    Array.from({ length: DEFAULT_ROWS }, () =>
      Array(DEFAULT_COLS).fill({ value: "", formula: "", style: {} })
    )
  );

  // Track row heights and column widths
  const [rowHeights, setRowHeights] = useState(Array(DEFAULT_ROWS).fill(25));
  const [colWidths, setColWidths] = useState(Array(DEFAULT_COLS).fill(100));

  const [selectedCell, setSelectedCell] = useState(null);
  const [selectionRange, setSelectionRange] = useState(null);
  const [formulaInput, setFormulaInput] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [dragStartPos, setDragStartPos] = useState(null);
  const [isResizing, setIsResizing] = useState(null); // null or {type: 'row'/'col', index: number}
  
  const formulaInputRef = useRef(null);
  const tableRef = useRef(null);
  
  const [activeStyles, setActiveStyles] = useState({
    bold: false,
    italic: false,
    underline: false,
    fontSize: "12px",
    textColor: "#000000",
    bgColor: "#ffffff",
    alignment: "left",
  });

  // Convert column index to letter (A, B, C...)
  const indexToColLetter = (idx) => {
    let letter = '';
    while (idx >= 0) {
      letter = String.fromCharCode(65 + (idx % 26)) + letter;
      idx = Math.floor(idx / 26) - 1;
    }
    return letter;
  };

  // Convert letter to column index (A->0, B->1, AA->26, etc.)
  const colLetterToIndex = (str) => {
    let result = 0;
    for (let i = 0; i < str.length; i++) {
      result = result * 26 + (str.charCodeAt(i) - 64);
    }
    return result - 1;
  };

  // Handle cell selection
  const handleCellSelect = (row, col, extend = false) => {
    if (extend && selectedCell) {
      // Extend selection range
      setSelectionRange({
        startRow: Math.min(selectedCell.row, row),
        startCol: Math.min(selectedCell.col, col),
        endRow: Math.max(selectedCell.row, row),
        endCol: Math.max(selectedCell.col, col)
      });
    } else {
      // New selection
      setSelectedCell({ row, col });
      setSelectionRange(null);
      setFormulaInput(cells[row][col]?.formula || cells[row][col]?.value || "");
      
      // Focus formula input
      if (formulaInputRef.current) {
        formulaInputRef.current.focus();
      }
      
      // Update active styles based on selected cell
      if (cells[row][col]?.style) {
        setActiveStyles({
          ...activeStyles,
          ...cells[row][col].style
        });
      } else {
        // Reset to defaults if no styles
        setActiveStyles({
          bold: false,
          italic: false,
          underline: false,
          fontSize: "12px",
          textColor: "#000000",
          bgColor: "#ffffff",
          alignment: "left",
        });
      }
    }
  };

  // Handle double click on cell (for direct editing)
  const handleCellDoubleClick = (row, col) => {
    setSelectedCell({ row, col, editing: true });
    if (formulaInputRef.current) {
      formulaInputRef.current.focus();
      formulaInputRef.current.select();
    }
  };

  // Handle drag start
  const handleDragStart = (e, row, col) => {
    if (e.nativeEvent.button !== 0) return; // Only left mouse button
    
    setIsDragging(true);
    setDragStartPos({ row, col });
    
    // Prevent text selection during drag
    document.body.classList.add('spreadsheet-dragging');
  };

  // Handle drag over a cell
  const handleDragOver = (row, col) => {
    if (!isDragging || !dragStartPos) return;
    
    handleCellSelect(dragStartPos.row, dragStartPos.col, true);
    setSelectionRange({
      startRow: Math.min(dragStartPos.row, row),
      startCol: Math.min(dragStartPos.col, col),
      endRow: Math.max(dragStartPos.row, row),
      endCol: Math.max(dragStartPos.col, col)
    });
  };

  // Handle drag end
  const handleDragEnd = () => {
    setIsDragging(false);
    setDragStartPos(null);
    document.body.classList.remove('spreadsheet-dragging');
  };

  // Handle resize start
  const handleResizeStart = (e, type, index) => {
    e.preventDefault();
    e.stopPropagation();
    setIsResizing({ type, index });
    
    // Add resize cursor and prevent text selection
    document.body.classList.add('spreadsheet-resizing');
    
    // Track initial mouse position
    const startX = e.clientX;
    const startY = e.clientY;
    
    const handleMouseMove = (moveEvent) => {
      if (type === 'col') {
        const delta = moveEvent.clientX - startX;
        setColWidths(prev => {
          const newWidths = [...prev];
          newWidths[index] = Math.max(50, prev[index] + delta);
          return newWidths;
        });
      } else if (type === 'row') {
        const delta = moveEvent.clientY - startY;
        setRowHeights(prev => {
          const newHeights = [...prev];
          newHeights[index] = Math.max(20, prev[index] + delta);
          return newHeights;
        });
      }
    };
    
    const handleMouseUp = () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
      setIsResizing(null);
      document.body.classList.remove('spreadsheet-resizing');
    };
    
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  // Handle formula input change
  const handleFormulaChange = (e) => {
    setFormulaInput(e.target.value);
  };

  // Apply formula to selected cell and update dependencies
  const applyFormula = () => {
    if (!selectedCell) return;
    
    const { row, col } = selectedCell;
    const isFormula = formulaInput.startsWith("=");
    
    setCells(prev => {
      const newCells = JSON.parse(JSON.stringify(prev));
      
      // Update the cell with the new formula
      newCells[row][col] = {
        ...newCells[row][col],
        value: isFormula ? evaluateFormula(formulaInput, newCells) : formulaInput,
        formula: isFormula ? formulaInput : "",
      };
      
      // Update all cells that depend on this one
      if (!isFormula) {
        updateDependentCells(newCells);
      }
      
      return newCells;
    });
  };
  
  // Update all cells with formulas that might depend on changed cells
  const updateDependentCells = (cellsData) => {
    let hasUpdates = true;
    let iterations = 0;
    const maxIterations = 100; // Prevent infinite loops
    
    // Keep updating until no more changes or max iterations reached
    while (hasUpdates && iterations < maxIterations) {
      hasUpdates = false;
      iterations++;
      
      for (let r = 0; r < cellsData.length; r++) {
        for (let c = 0; c < cellsData[r].length; c++) {
          const cell = cellsData[r][c];
          
          if (cell.formula && cell.formula.startsWith("=")) {
            const newValue = evaluateFormula(cell.formula, cellsData);
            
            if (newValue !== cell.value) {
              cellsData[r][c] = {
                ...cell,
                value: newValue
              };
              hasUpdates = true;
            }
          }
        }
      }
    }
    
    return cellsData;
  };
  
  // Handle keydown in formula input
  const handleFormulaKeyDown = (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      applyFormula();
      
      // Move to the next cell down after Enter
      if (selectedCell && selectedCell.row < rowCount - 1) {
        const nextRow = selectedCell.row + 1;
        handleCellSelect(nextRow, selectedCell.col);
      }
    } else if (e.key === "Tab") {
      e.preventDefault();
      applyFormula();
      
      // Move to the next cell right after Tab
      if (selectedCell && selectedCell.col < colCount - 1) {
        const nextCol = selectedCell.col + 1;
        handleCellSelect(selectedCell.row, nextCol);
      }
    } else if (e.key === "Escape") {
      // Cancel editing and revert to original value
      if (selectedCell) {
        const { row, col } = selectedCell;
        setFormulaInput(cells[row][col]?.formula || cells[row][col]?.value || "");
      }
    }
  };

  // Apply style to selected cells
  const applyStyle = (styleProperty, value) => {
    if (!selectedCell && !selectionRange) return;
    
    // Update active styles state
    setActiveStyles(prev => ({
      ...prev,
      [styleProperty]: value
    }));
    
    setCells(prev => {
      const newCells = JSON.parse(JSON.stringify(prev));
      
      // Determine the range of cells to update
      let startRow, endRow, startCol, endCol;
      
      if (selectionRange) {
        startRow = selectionRange.startRow;
        endRow = selectionRange.endRow;
        startCol = selectionRange.startCol;
        endCol = selectionRange.endCol;
      } else {
        startRow = endRow = selectedCell.row;
        startCol = endCol = selectedCell.col;
      }
      
      // Apply style to all cells in the range
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          newCells[r][c] = {
            ...newCells[r][c],
            style: {
              ...newCells[r][c].style,
              [styleProperty]: value
            }
          };
        }
      }
      
      return newCells;
    });
  };

  // Toggle style (bold, italic, underline)
  const toggleStyle = (styleProperty) => {
    applyStyle(styleProperty, !activeStyles[styleProperty]);
  };

  // Handle function selection
  const handleFunctionSelect = (e) => {
    const func = e.target.value;
    if (!func) return;
    
    setFormulaInput(prev => {
      const newValue = `=${func}()`;
      // Reset the select
      e.target.value = "";
      return newValue;
    });
    
    if (formulaInputRef.current) {
      formulaInputRef.current.focus();
      // Position cursor between parentheses
      setTimeout(() => {
        formulaInputRef.current.setSelectionRange(
          func.length + 2,
          func.length + 2
        );
      }, 0);
    }
  };

  // Add a new row
  const addRow = (afterIndex = rowCount - 1) => {
    const newRowCount = rowCount + 1;
    setRowCount(newRowCount);
    setRowHeights(prev => {
      const newHeights = [...prev];
      newHeights.splice(afterIndex + 1, 0, 25); // Default height
      return newHeights;
    });
    
    setCells(prev => {
      const newCells = [...prev];
      const newRow = Array(colCount).fill({ value: "", formula: "", style: {} });
      newCells.splice(afterIndex + 1, 0, newRow);
      return newCells;
    });
  };

  // Add a new column
  const addColumn = (afterIndex = colCount - 1) => {
    const newColCount = colCount + 1;
    setColCount(newColCount);
    setColWidths(prev => {
      const newWidths = [...prev];
      newWidths.splice(afterIndex + 1, 0, 100); // Default width
      return newWidths;
    });
    
    setCells(prev => {
      const newCells = prev.map(row => {
        const newRow = [...row];
        newRow.splice(afterIndex + 1, 0, { value: "", formula: "", style: {} });
        return newRow;
      });
      return newCells;
    });
  };

  // Delete a row
  const deleteRow = (index) => {
    if (rowCount <= 1) return;
    
    const newRowCount = rowCount - 1;
    setRowCount(newRowCount);
    setRowHeights(prev => {
      const newHeights = [...prev];
      newHeights.splice(index, 1);
      return newHeights;
    });
    
    setCells(prev => {
      const newCells = [...prev];
      newCells.splice(index, 1);
      return updateDependentCells(newCells);
    });
  };

  // Delete a column
  const deleteColumn = (index) => {
    if (colCount <= 1) return;
    
    const newColCount = colCount - 1;
    setColCount(newColCount);
    setColWidths(prev => {
      const newWidths = [...prev];
      newWidths.splice(index, 1);
      return newWidths;
    });
    
    setCells(prev => {
      const newCells = prev.map(row => {
        const newRow = [...row];
        newRow.splice(index, 1);
        return newRow;
      });
      return updateDependentCells(newCells);
    });
  };

  // Show context menu for row/column header
  const showHeaderContextMenu = (e, type, index) => {
    e.preventDefault();
    
    // Create context menu
    const menu = document.createElement('div');
    menu.className = 'spreadsheet-context-menu';
    menu.style.top = `${e.clientY}px`;
    menu.style.left = `${e.clientX}px`;
    
    // Add menu items
    const insertItem = document.createElement('div');
    insertItem.className = 'context-menu-item';
    insertItem.textContent = `Insert ${type}`;
    insertItem.onclick = () => {
      if (type === 'row') {
        addRow(index);
      } else {
        addColumn(index);
      }
      document.body.removeChild(menu);
    };
    
    const deleteItem = document.createElement('div');
    deleteItem.className = 'context-menu-item';
    deleteItem.textContent = `Delete ${type}`;
    deleteItem.onclick = () => {
      if (type === 'row') {
        deleteRow(index);
      } else {
        deleteColumn(index);
      }
      document.body.removeChild(menu);
    };
    
    menu.appendChild(insertItem);
    menu.appendChild(deleteItem);
    document.body.appendChild(menu);
    
    // Remove menu when clicking elsewhere
    const removeMenu = (e) => {
      if (!menu.contains(e.target)) {
        document.body.removeChild(menu);
        document.removeEventListener('click', removeMenu);
      }
    };
    
    setTimeout(() => {
      document.addEventListener('click', removeMenu);
    }, 0);
  };

  // Basic formula evaluation with cell dependencies
  const evaluateFormula = (formula, cellsData) => {
    if (!formula.startsWith("=")) return formula;
    
    // Remove the equals sign
    const expression = formula.substring(1);
    
    // Handle basic SUM function
    if (expression.startsWith("SUM(") && expression.endsWith(")")) {
      const range = expression.substring(4, expression.length - 1);
      const [start, end] = range.split(":");
      
      if (start && end) {
        try {
          // Parse cell references like A1:B3
          const startColStr = start.replace(/[0-9]/g, '');
          const startRowStr = start.replace(/[^0-9]/g, '');
          
          const endColStr = end.replace(/[0-9]/g, '');
          const endRowStr = end.replace(/[^0-9]/g, '');
          
          const startCol = colLetterToIndex(startColStr);
          const startRow = parseInt(startRowStr) - 1;
          const endCol = colLetterToIndex(endColStr);
          const endRow = parseInt(endRowStr) - 1;
          
          let sum = 0;
          for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
              if (r >= 0 && r < cellsData.length && c >= 0 && c < cellsData[0].length) {
                const cellValue = cellsData[r][c].value;
                if (!isNaN(parseFloat(cellValue))) {
                  sum += parseFloat(cellValue);
                }
              }
            }
          }
          return sum.toString();
        } catch (e) {
          return "#ERROR!";
        }
      }
    }
    
    // Handle AVERAGE function
    if (expression.startsWith("AVERAGE(") && expression.endsWith(")")) {
      const range = expression.substring(8, expression.length - 1);
      const [start, end] = range.split(":");
      
      if (start && end) {
        try {
          const startColStr = start.replace(/[0-9]/g, '');
          const startRowStr = start.replace(/[^0-9]/g, '');
          
          const endColStr = end.replace(/[0-9]/g, '');
          const endRowStr = end.replace(/[^0-9]/g, '');
          
          const startCol = colLetterToIndex(startColStr);
          const startRow = parseInt(startRowStr) - 1;
          const endCol = colLetterToIndex(endColStr);
          const endRow = parseInt(endRowStr) - 1;
          
          let sum = 0;
          let count = 0;
          
          for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
              if (r >= 0 && r < cellsData.length && c >= 0 && c < cellsData[0].length) {
                const cellValue = cellsData[r][c].value;
                if (!isNaN(parseFloat(cellValue))) {
                  sum += parseFloat(cellValue);
                  count++;
                }
              }
            }
          }
          
          return count > 0 ? (sum / count).toFixed(2) : "#DIV/0!";
        } catch (e) {
          return "#ERROR!";
        }
      }
    }
    
    // Basic arithmetic evaluation
    try {
      // Replace cell references (like A1) with their values
      let processedExpression = expression.replace(/[A-Z]+[0-9]+/g, (match) => {
        const colStr = match.replace(/[0-9]/g, '');
        const rowStr = match.replace(/[^0-9]/g, '');
        
        const col = colLetterToIndex(colStr);
        const row = parseInt(rowStr) - 1;
        
        if (row >= 0 && row < cellsData.length && col >= 0 && col < cellsData[0].length) {
          const cellValue = cellsData[row][col].value;
          return !isNaN(parseFloat(cellValue)) ? cellValue : "0";
        }
        return "0";
      });
      
      // Use Function constructor instead of eval for slightly safer execution
      const result = new Function(`return ${processedExpression}`)();
      return result.toString();
    } catch (e) {
      return "#ERROR!";
    }
  };

  // Check if a cell is in the selection range
  const isCellInSelection = (row, col) => {
    if (!selectionRange) return false;
    
    return (
      row >= selectionRange.startRow &&
      row <= selectionRange.endRow &&
      col >= selectionRange.startCol &&
      col <= selectionRange.endCol
    );
  };

  useEffect(() => {
    // Set up event listeners for drag and drop
    const handleMouseUp = () => {
      if (isDragging) {
        handleDragEnd();
      }
    };
    
    document.addEventListener('mouseup', handleMouseUp);
    
    return () => {
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isDragging]);

  return (
    <div className="spreadsheet-container">
      {/* Main Toolbar */}
      <div className="toolbar">
        <div className="toolbar-section">
          <button 
            className={`toolbar-button ${activeStyles.bold ? 'active' : ''}`}
            onClick={() => toggleStyle('bold')}
            title="Bold"
          >
            B
          </button>
          <button 
            className={`toolbar-button ${activeStyles.italic ? 'active' : ''}`}
            onClick={() => toggleStyle('italic')}
            title="Italic"
          >
            I
          </button>
          <button 
            className={`toolbar-button ${activeStyles.underline ? 'active' : ''}`}
            onClick={() => toggleStyle('underline')}
            title="Underline"
          >
            U
          </button>
        </div>

        <div className="toolbar-section">
          <select 
            className="toolbar-select"
            value={activeStyles.fontSize}
            onChange={(e) => applyStyle('fontSize', e.target.value)}
            title="Font Size"
          >
            <option value="10px">10</option>
            <option value="12px">12</option>
            <option value="14px">14</option>
            <option value="18px">18</option>
            <option value="24px">24</option>
          </select>
        </div>

        <div className="toolbar-section">
          <div className="color-picker-container">
            <span>Text:</span>
            <input 
              type="color" 
              className="toolbar-color-picker"
              value={activeStyles.textColor}
              onChange={(e) => applyStyle('textColor', e.target.value)}
              title="Text Color"
            />
          </div>
          <div className="color-picker-container">
            <span>Fill:</span>
            <input 
              type="color" 
              className="toolbar-color-picker"
              value={activeStyles.bgColor}
              onChange={(e) => applyStyle('bgColor', e.target.value)}
              title="Background Color"
            />
          </div>
        </div>

        <div className="toolbar-section">
          <button 
            className={`toolbar-button ${activeStyles.alignment === 'left' ? 'active' : ''}`}
            onClick={() => applyStyle('alignment', 'left')}
            title="Align Left"
          >
            ←
          </button>
          <button 
            className={`toolbar-button ${activeStyles.alignment === 'center' ? 'active' : ''}`}
            onClick={() => applyStyle('alignment', 'center')}
            title="Align Center"
          >
            ↔
          </button>
          <button 
            className={`toolbar-button ${activeStyles.alignment === 'right' ? 'active' : ''}`}
            onClick={() => applyStyle('alignment', 'right')}
            title="Align Right"
          >
            →
          </button>
        </div>

        <div className="toolbar-section">
          <select 
            className="toolbar-select function-select" 
            title="Functions"
            onChange={handleFunctionSelect}
            value=""
          >
            <option value="">Functions</option>
            <option value="SUM">SUM</option>
            <option value="AVERAGE">AVERAGE</option>
            <option value="COUNT">COUNT</option>
            <option value="MAX">MAX</option>
            <option value="MIN">MIN</option>
          </select>
        </div>
        
        <div className="toolbar-section actions">
          <button 
            className="toolbar-button" 
            title="Add Row"
            onClick={() => addRow(rowCount - 1)}
          >
            + Row
          </button>
          <button 
            className="toolbar-button" 
            title="Add Column"
            onClick={() => addColumn(colCount - 1)}
          >
            + Col
          </button>
        </div>
      </div>

      {/* Current Cell Reference */}
      <div className="cell-reference">
        {selectedCell 
          ? `${indexToColLetter(selectedCell.col)}${selectedCell.row + 1}` 
          : selectionRange 
            ? `${indexToColLetter(selectionRange.startCol)}${selectionRange.startRow + 1}:${indexToColLetter(selectionRange.endCol)}${selectionRange.endRow + 1}`
            : ""
        }
      </div>

      {/* Formula Bar */}
      <div className="formula-bar">
        <span className="formula-icon">fx</span>
        <input
          ref={formulaInputRef}
          type="text"
          className="formula-input"
          value={formulaInput}
          onChange={handleFormulaChange}
          onBlur={applyFormula}
          onKeyDown={handleFormulaKeyDown}
          placeholder="Enter a value or formula starting with ="
        />
      </div>

      {/* Spreadsheet Grid */}
      <div className="table-container">
        <table className="spreadsheet-table" ref={tableRef}>
          <colgroup>
            <col style={{ width: "40px" }} /> {/* Corner header */}
            {Array.from({ length: colCount }).map((_, i) => (
              <col key={i} style={{ width: `${colWidths[i]}px` }} />
            ))}
          </colgroup>
          <thead>
            <tr>
              <th className="corner-header"></th>
              {Array.from({ length: colCount }).map((_, i) => (
                <th 
                  key={i} 
                  className="column-header"
                  onContextMenu={(e) => showHeaderContextMenu(e, 'column', i)}
                >
                  {indexToColLetter(i)}
                  <div 
                    className="resize-handle col-resize"
                    onMouseDown={(e) => handleResizeStart(e, 'col', i)}
                  ></div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: rowCount }).map((_, rowIndex) => (
              <tr key={rowIndex} style={{ height: `${rowHeights[rowIndex]}px` }}>
                <th 
                  className="row-header"
                  onContextMenu={(e) => showHeaderContextMenu(e, 'row', rowIndex)}
                >
                  {rowIndex + 1}
                  <div 
                    className="resize-handle row-resize"
                    onMouseDown={(e) => handleResizeStart(e, 'row', rowIndex)}
                  ></div>
                </th>
                {Array.from({ length: colCount }).map((_, colIndex) => {
                  const cell = cells[rowIndex]?.[colIndex] || { value: "", style: {} };
                  const isSelected = selectedCell?.row === rowIndex && selectedCell?.col === colIndex;
                  const isInSelection = isCellInSelection(rowIndex, colIndex);
                  
                  return (
                    <Cell
                      key={`cell-${rowIndex}-${colIndex}`}
                      row={rowIndex}
                      col={colIndex}
                      value={cell.value}
                      style={cell.style}
                      isSelected={isSelected}
                      isInSelection={isInSelection}
                      onSelect={handleCellSelect}
                      onDoubleClick={handleCellDoubleClick}
                      onDragStart={handleDragStart}
                      onDragOver={handleDragOver}
                    />
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default Spreadsheet;