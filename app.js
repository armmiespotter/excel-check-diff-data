const xlsx = require("xlsx");
const diff = require("json-diff");

// Load the Excel file
const workbook = xlsx.readFile("data.xlsx");

// Get the first and second sheets
const beforeSheetName = workbook.SheetNames[0];
const afterSheetName = workbook.SheetNames[1];

// Convert sheets to JSON
const beforeData = xlsx.utils.sheet_to_json(workbook.Sheets[beforeSheetName]);
const afterData = xlsx.utils.sheet_to_json(workbook.Sheets[afterSheetName]);

// Function to find a matching row in the after data
const findMatchingRow = (row, data, keyColumns) => {
  return data.find((d) => keyColumns.every((key) => d[key] === row[key]));
};

// Function to compare data
const compareData = (before, after, keyColumns) => {
  const differences = [];

  before.forEach((beforeRow) => {
    const afterRow = findMatchingRow(beforeRow, after, keyColumns);
    if (!afterRow) {
      differences.push({
        ...beforeRow,
        Differences: "Not found in after data",
      });
    } else {
      const diffResult = diff.diff(beforeRow, afterRow);
      if (diffResult) {
        differences.push({
          ...beforeRow,
          Differences: JSON.stringify(diffResult),
        });
      } else {
        differences.push({ ...beforeRow, Differences: "No differences" });
      }
    }
  });

  // Check for new rows in the after data that are not in the before data
  after.forEach((afterRow) => {
    const beforeRow = findMatchingRow(afterRow, before, keyColumns);
    if (!beforeRow) {
      differences.push({
        ...afterRow,
        Differences: "Not found in before data",
      });
    }
  });

  return differences;
};

// Define key columns to match rows
const keyColumns = ["Name", "Group"];

// Compare the data
const differences = compareData(beforeData, afterData, keyColumns);

// Save differences to a new Excel file
const newWorkbook = xlsx.utils.book_new();
const newSheet = xlsx.utils.json_to_sheet(differences);
xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Differences");
xlsx.writeFile(newWorkbook, "comparison_result.xlsx");

console.log('Comparison complete. Check "comparison_result.xlsx" for details.');
