const spreadsheet = document.getElementById("spreadsheet");
let activeCell = null;

// Create spreadsheet
function createSpreadsheet(rows, cols) {
  for (let i = 0; i <= rows; i++) {
    const row = spreadsheet.insertRow();
    for (let j = 0; j <= cols; j++) {
      const cell = row.insertCell();
      if (i === 0 && j > 0) {
        cell.innerText = String.fromCharCode(64 + j); // Column headers (A, B, C...)
      } else if (j === 0 && i > 0) {
        cell.innerText = i; // Row headers (1, 2, 3...)
      } else if (i > 0 && j > 0) {
        cell.contentEditable = true;

        cell.addEventListener("click", () => {
          if (activeCell) activeCell.style.outline = "none";
          activeCell = cell;
          cell.style.outline = "2px solid #007BFF";
        });
      }
    }
  }
  enableDrag();
  enableColumnResize();
  enableRowResize();
}

// Drag-and-Drop Functionality
function enableDrag() {
  const cells = document.querySelectorAll("td:not(:first-child):not([draggable=false])");

  cells.forEach((cell) => {
    cell.setAttribute("draggable", true);

    cell.addEventListener("dragstart", (e) => {
      e.dataTransfer.setData("text/plain", cell.innerText);
      e.target.style.opacity = "0.5"; // Feedback for dragging
    });

    cell.addEventListener("dragend", (e) => {
      e.target.style.opacity = "1";
    });

    cell.addEventListener("dragover", (e) => e.preventDefault());

    cell.addEventListener("drop", (e) => {
      e.preventDefault();
      const data = e.dataTransfer.getData("text/plain");
      const targetCell = e.target;

      if (targetCell.tagName === "TD") {
        targetCell.innerText = data;
        const sourceCell = document.querySelector("td[style*='opacity: 0.5']");
        if (sourceCell) sourceCell.innerText = ""; // Clear the source cell (move functionality)
      }
    });
  });
}

// Formatting
function formatCell(style) {
  if (!activeCell) return alert("Select a cell first!");
  if (style === "bold") activeCell.style.fontWeight = activeCell.style.fontWeight === "bold" ? "normal" : "bold";
  if (style === "italic") activeCell.style.fontStyle = activeCell.style.fontStyle === "italic" ? "normal" : "italic";
}

function changeTextColor() {
  if (!activeCell) return alert("Select a cell first!");
  const color = prompt("Enter text color:", "black");
  if (color) activeCell.style.color = color;
}

// Add/Delete Rows and Columns
function addRow() {
  const row = spreadsheet.insertRow();
  for (let i = 0; i < spreadsheet.rows[0].cells.length; i++) {
    const cell = row.insertCell();
    cell.contentEditable = true;
  }
}

function addColumn() {
  for (let i = 0; i < spreadsheet.rows.length; i++) {
    const cell = spreadsheet.rows[i].insertCell();
    if (i > 0) cell.contentEditable = true;
  }
}

function deleteRow() {
  if (spreadsheet.rows.length > 1) spreadsheet.deleteRow(-1);
}

function deleteColumn() {
  if (spreadsheet.rows[0].cells.length > 1) {
    for (let i = 0; i < spreadsheet.rows.length; i++) {
      spreadsheet.rows[i].deleteCell(-1);
    }
  }
}

// Dynamic Resizing
function enableColumnResize() {
  const headerCells = spreadsheet.rows[0].cells;

  for (let i = 1; i < headerCells.length; i++) {
    const cell = headerCells[i];
    cell.style.position = "relative";

    const resizeHandle = document.createElement("div");
    resizeHandle.style.position = "absolute";
    resizeHandle.style.right = "0";
    resizeHandle.style.top = "0";
    resizeHandle.style.width = "5px";
    resizeHandle.style.height = "100%";
    resizeHandle.style.cursor = "col-resize";

    cell.appendChild(resizeHandle);

    resizeHandle.addEventListener("mousedown", (e) => {
      const startX = e.pageX;
      const startWidth = cell.offsetWidth;

      const mouseMoveHandler = (event) => {
        const newWidth = startWidth + (event.pageX - startX);
        for (let row of spreadsheet.rows) {
          row.cells[i].style.width = `${newWidth}px`;
        }
      };

      const mouseUpHandler = () => {
        document.removeEventListener("mousemove", mouseMoveHandler);
        document.removeEventListener("mouseup", mouseUpHandler);
      };

      document.addEventListener("mousemove", mouseMoveHandler);
      document.addEventListener("mouseup", mouseUpHandler);
    });
  }
}

function enableRowResize() {
  const rows = spreadsheet.rows;

  for (let i = 1; i < rows.length; i++) {
    const rowHeader = rows[i].cells[0];
    rowHeader.style.position = "relative";

    const resizeHandle = document.createElement("div");
    resizeHandle.style.position = "absolute";
    resizeHandle.style.bottom = "0";
    resizeHandle.style.left = "0";
    resizeHandle.style.width = "100%";
    resizeHandle.style.height = "5px";
    resizeHandle.style.cursor = "row-resize";

    rowHeader.appendChild(resizeHandle);

    resizeHandle.addEventListener("mousedown", (e) => {
      const startY = e.pageY;
      const startHeight = rowHeader.parentElement.offsetHeight;

      const mouseMoveHandler = (event) => {
        const newHeight = startHeight + (event.pageY - startY);
        rows[i].style.height = `${newHeight}px`;
      };

      const mouseUpHandler = () => {
        document.removeEventListener("mousemove", mouseMoveHandler);
        document.removeEventListener("mouseup", mouseUpHandler);
      };

      document.addEventListener("mousemove", mouseMoveHandler);
      document.addEventListener("mouseup", mouseUpHandler);
    });
  }
}

// Mathematical Functions
function applyFormula() {
  const formulaInput = document.getElementById("formula-input").value.trim().toUpperCase();
  if (!activeCell) return alert("Select a cell to apply the formula!");
  if (!formulaInput.startsWith("=")) return alert("Formula must start with '='!");

  const formulaParts = formulaInput.slice(1).split("(");
  const functionName = formulaParts[0];
  const range = formulaParts[1]?.replace(")", "").trim();

  if (!range) return alert("Invalid formula. Example: =SUM(A1:A3)");

  const rangeRegex = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/;
  const match = range.match(rangeRegex);

  if (!match) return alert("Invalid range format. Use format like A1:A3.");

  const [, startCol, startRow, endCol, endRow] = match;
  const startColIndex = startCol.charCodeAt(0) - 64;
  const endColIndex = endCol.charCodeAt(0) - 64;
  const startRowIndex = parseInt(startRow, 10);
  const endRowIndex = parseInt(endRow, 10);

  const values = [];
  for (let i = startRowIndex; i <= endRowIndex; i++) {
    for (let j = startColIndex; j <= endColIndex; j++) {
      const cell = spreadsheet.rows[i]?.cells[j];
      const value = parseFloat(cell?.innerText) || 0;
      values.push(value);
    }
  }

  let result;
  switch (functionName) {
    case "SUM":
      result = values.reduce((a, b) => a + b, 0);
      break;
    case "AVERAGE":
      result = values.reduce((a, b) => a + b, 0) / values.length;
      break;
    case "MAX":
      result = Math.max(...values);
      break;
    case "MIN":
      result = Math.min(...values);
      break;
    case "COUNT":
      result = values.filter((v) => !isNaN(v)).length;
      break;
    case "MEDIAN":
      values.sort((a, b) => a - b);
      const mid = Math.floor(values.length / 2);
      result = values.length % 2 !== 0 ? values[mid] : (values[mid - 1] + values[mid]) / 2;
      break;      
    case "PRODUCT":
      result = values.reduce((a, b) => a * b, 1);
      break;
      
    default:
      return alert("Unsupported function. Use SUM, AVERAGE, MAX, MIN, or COUNT.");
  }

  activeCell.innerText = result;
}

// Data Quality Functions
function trimCell() {
  if (!activeCell) return alert("Select a cell first!");
  activeCell.innerText = activeCell.innerText.trim();
}

function convertToUpper() {
  if (!activeCell) return alert("Select a cell first!");
  activeCell.innerText = activeCell.innerText.toUpperCase();
}

function convertToLower() {
  if (!activeCell) return alert("Select a cell first!");
  activeCell.innerText = activeCell.innerText.toLowerCase();
}

function convertToTitleCase() {
  if (!activeCell) return alert("Select a cell first!");
  activeCell.innerText = activeCell.innerText
    .toLowerCase()
    .split(" ")
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

function reverseText() {
  if (!activeCell) return alert("Select a cell first!");
  activeCell.innerText = activeCell.innerText.split("").reverse().join("");
}


function removeDuplicates() {
  const rangeInput = prompt("Enter range to remove duplicates (e.g., A1:D5):", "");
  if (!rangeInput) return;

  const rangeRegex = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/;
  const match = rangeInput.match(rangeRegex);

  if (!match) {
    alert("Invalid range format. Use format like A1:D5.");
    return;
  }

  const [, startCol, startRow, endCol, endRow] = match;
  const startColIndex = startCol.charCodeAt(0) - 64;
  const endColIndex = endCol.charCodeAt(0) - 64;
  const startRowIndex = parseInt(startRow, 10);
  const endRowIndex = parseInt(endRow, 10);

  const seenRows = new Set();
  for (let i = startRowIndex; i <= endRowIndex; i++) {
    const rowValues = [];
    for (let j = startColIndex; j <= endColIndex; j++) {
      const cell = spreadsheet.rows[i]?.cells[j];
      rowValues.push(cell?.innerText || "");
    }
    const rowKey = rowValues.join("|");
    if (seenRows.has(rowKey)) {
      spreadsheet.deleteRow(i);
      i--; // Adjust index after deletion
    } else {
      seenRows.add(rowKey);
    }
  }
}

function findAndReplace() {
  const rangeInput = prompt("Enter range for find and replace (e.g., A1:D5):", "");
  if (!rangeInput) return;

  const searchText = prompt("Enter text to find:", "");
  const replaceText = prompt("Enter replacement text:", "");

  if (!searchText || replaceText === null) return;

  const rangeRegex = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/;
  const match = rangeInput.match(rangeRegex);

  if (!match) {
    alert("Invalid range format. Use format like A1:D5.");
    return;
  }

  const [, startCol, startRow, endCol, endRow] = match;
  const startColIndex = startCol.charCodeAt(0) - 64;
  const endColIndex = endCol.charCodeAt(0) - 64;
  const startRowIndex = parseInt(startRow, 10);
    const endRowIndex = parseInt(endRow, 10);
  
    for (let i = startRowIndex; i <= endRowIndex; i++) {
      for (let j = startColIndex; j <= endColIndex; j++) {
        const cell = spreadsheet.rows[i]?.cells[j];
        if (cell && cell.innerText.includes(searchText)) {
          cell.innerText = cell.innerText.replace(new RegExp(searchText, "g"), replaceText);
        }
      }
    }
  }
 
// Initialize spreadsheet
createSpreadsheet(10, 10);



























  
  
  