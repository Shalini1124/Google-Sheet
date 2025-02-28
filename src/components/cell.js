import React from "react";

const Cell = ({ 
  row, 
  col, 
  value, 
  style = {}, 
  isSelected, 
  onSelect,
  onDoubleClick 
}) => {
  const handleClick = (e) => {
    onSelect(row, col);
    e.stopPropagation();
  };

  const handleDoubleClick = (e) => {
    if (onDoubleClick) {
      onDoubleClick(row, col);
    }
    e.stopPropagation();
  };

  
  const cellStyle = {
    fontWeight: style.bold ? "bold" : "normal",
    fontStyle: style.italic ? "italic" : "normal",
    textDecoration: style.underline ? "underline" : "none",
    fontSize: style.fontSize || "12px",
    color: style.textColor || "#000000",
    backgroundColor: style.bgColor || "#ffffff",
    textAlign: style.alignment || "left",
  };

  return (
    <td
      className={`spreadsheet-cell ${isSelected ? "selected" : ""}`}
      onClick={handleClick}
      onDoubleClick={handleDoubleClick}
      style={cellStyle}
    >
      {value}
    </td>
  );
};

export default Cell;