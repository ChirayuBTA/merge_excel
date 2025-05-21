const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

// Path to your folder containing Excel files
const folderPath = "./input_excels"; // <-- change this
const outputFolder = "./output";

// Read all Excel files in the folder
const files = fs
  .readdirSync(folderPath)
  .filter((file) => file.endsWith(".xlsx") || file.endsWith(".xls"));

// Store all rows here
let mergedData = [];

// Loop through each file
files.forEach((file) => {
  const filePath = path.join(folderPath, file);
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0]; // Read the first sheet
  const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

  mergedData = mergedData.concat(sheetData);
});

// Create a new workbook and add the merged data
const newWorkbook = xlsx.utils.book_new();
const newWorksheet = xlsx.utils.json_to_sheet(mergedData);
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "MergedData");

// Save the merged Excel file
const outputFilePath = path.join(outputFolder, "merged_output.xlsx");
xlsx.writeFile(newWorkbook, outputFilePath);

console.log("Merge complete. Saved as merged_output.xlsx");
