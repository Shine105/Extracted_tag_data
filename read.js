const xlsx = require('xlsx');

// Path to the input Excel file
const inputFilePath = 'AHERI_220_14_11_2023.xls';

// Read the Excel file
const workbook = xlsx.readFile(inputFilePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Function to extract all string values from a specified row, excluding values containing "DUMMY"
function extractRowValues(worksheet, rowNumber) {
  const rowValues = [];
  const range = xlsx.utils.decode_range(worksheet['!ref']);

  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellAddress = xlsx.utils.encode_cell({ c: col, r: rowNumber });
    const cell = worksheet[cellAddress];
    const value = cell ? cell.v : null; // Use null for empty cells

    // Check if the value is a string and does not contain "DUMMY"
    if (value !== null && typeof value === 'string' && !value.includes('DUMMY')) {
      rowValues.push({ value, col });
    }
  }

  return rowValues;
}

// Function to extract data from a specific column starting from a given row for a certain number of rows
function extractColumnData(worksheet, col, startRow, numRows) {
  const columnData = [];
  for (let row = startRow; row < startRow + numRows; row++) {
    const cellAddress = xlsx.utils.encode_cell({ c: col, r: row });
    const cell = worksheet[cellAddress];
    const value = cell ? cell.v : 'N/A'; // Use 'N/A' for empty cells
    columnData.push(value);
  }
  return columnData;
}

// Extract string values from the third row (row index 2, as it's 0-based)
const thirdRowValues = extractRowValues(worksheet, 2);

// Log the extracted values
console.log('String values in the third row (excluding values containing "DUMMY"):', thirdRowValues);

// Prepare data for the new worksheet
const outputData = [['SCADA Tag', ...Array.from({ length: 1440 }, (_, i) => `Data ${i + 1}`)]];

// Extract data for each SCADA tag and add to the outputData
thirdRowValues.forEach(({ value, col }) => {
  const columnData = extractColumnData(worksheet, col, 5, 1440); // Extract data from 6th row (index 5) and next 1440 rows
  outputData.push([value, ...columnData]);
});

// Create a new workbook and worksheet for the output data
const outputWorkbook = xlsx.utils.book_new();
const outputWorksheet = xlsx.utils.aoa_to_sheet(outputData);
xlsx.utils.book_append_sheet(outputWorkbook, outputWorksheet, 'Sheet1');

// Save the output workbook to an Excel file
const outputFilePath = 'Extracted_SCADA_Tag_Data.xlsx';
xlsx.writeFile(outputWorkbook, outputFilePath);

console.log('SCADA Tag data has been successfully written to Extracted_SCADA_Tag_Data.xlsx');
