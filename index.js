// Import required libraries
const express = require('express');
const ExcelJS = require('exceljs');
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 10000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Store items in memory
let items = [];

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'portal.html'));
});

app.get('/page.html',(req,res)=>{
  items=[];
  res.sendFile(path.join(__dirname, 'page.html'));
});
app.post('/addItem', (req, res) => {
  const { item, quantity, unitPrice } = req.body;

  if (!item || !quantity || !unitPrice) {
    res.status(400).json({ error: 'All fields are required' });
  } else {
    const totalAmount = quantity * unitPrice;
    items.push({ item, quantity, unitPrice, totalAmount });
    res.status(201).json({ message: 'Item added successfully' });
  }
});

app.get('/getItems', (req, res) => {
  res.json(items);
});

app.post('/exportToExcel', async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('BOQ');
  
  const startRow = 3;  // Start from the 3rd row
  const startColumn = 2; // Start from the 2nd column (B)

  // Add column headers and apply border styling
  const headerRow = worksheet.getRow(startRow);
  headerRow.getCell(startColumn).value = 'Item';
  headerRow.getCell(startColumn + 1).value = 'Quantity';
  headerRow.getCell(startColumn + 2).value = 'Unit Price';
  
  // Merge the header cell for "Total Amount" column
  worksheet.mergeCells(startRow, startColumn + 3, startRow, startColumn + 4);
  headerRow.getCell(startColumn + 3).value = 'Total Amount';

  // Apply border styling to the header row
  headerRow.eachCell((cell) => {
    cell.border = {
      top: { style: 'thick' },
      left: { style: 'thick' },
      bottom: { style: 'thick' },
      right: { style: 'thick' },
    };
  });

  items.forEach((item, index) => {
    let row = worksheet.getRow(startRow + index + 1);
    if (!row) {
      row = worksheet.addRow([]);
    }
    row.getCell(startColumn).value = item.item;
    row.getCell(startColumn + 1).value = item.quantity;
    row.getCell(startColumn + 2).value = item.unitPrice;
    row.getCell(startColumn + 3).value = item.totalAmount;

    // Merge cells for the Total Amount column (columns 4 and 5 in 1-based index)
    worksheet.mergeCells(startRow + index + 1, startColumn + 3, startRow + index + 1, startColumn + 4);

    // Apply border styling to the data rows
    row.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        right: { style: 'thick' },
      };
    });

    // Check if it's the last row and set the bottom border to 'thick' for all cells in the last row
    if (index === items.length - 1) {
      row.eachCell((cell) => {
        cell.border.bottom = { style: 'thick' };
      });
    }
  });
  
  // Generate the Excel file
  const excelBuffer = await workbook.xlsx.writeBuffer();
  
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=BOQ.xlsx');
  res.send(excelBuffer);
  items=[];
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
