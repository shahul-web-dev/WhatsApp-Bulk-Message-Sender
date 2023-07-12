/** @format */

import XlsxPopulate from "xlsx-populate";

// Function to add data to the Excel file
async function addDataToExcel(variableData) {
  try {
    // Load the Excel file
    const workbook = await XlsxPopulate.fromFileAsync(
      ".\\NotSent.xlsx"
    );

    // Select the first sheet
    const sheet = workbook.sheet(0);

    // Find the last used row in column A
    const lastRow = sheet.usedRange().endCell("down").rowNumber();

    // Add the variable data to the next row
    sheet.cell(`A${lastRow + 1}`).value(variableData);

    // Save the modified workbook
    await workbook.toFileAsync(".\\NotSent.xlsx");
    console.log("Data added successfully to the Excel file.");
  } catch (error) {
    console.error("An error occurred:", error);
  }
}

// Usage
const variableData = "+919042";
addDataToExcel(variableData);
