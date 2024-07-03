import * as ExcelJS from "exceljs";
import { Cell, Workbook, Worksheet } from "exceljs";

export const generateExcel = async () => {
  const workbook: Workbook = new ExcelJS.Workbook();
  const worksheet: Worksheet = workbook.addWorksheet("BdC", {
    views: [{ showGridLines: false }], // Hide grid lines
  });

  // Set cell values
  const cellA1: Cell = worksheet.getCell("A1");

  // Define cell values
  cellA1.value = "Excel example";

  // Build here your excel file

  try {
    // Save the workbook to a buffer
    const buffer = await workbook.xlsx.writeBuffer();

    // Save file in "./" as "ExcelFile.xlsx"
    // const file = await workbook.xlsx.writeFile("./ExcelFile.xlsx");

    // We return the buffer, and we will send it to our sharepoint api
    return buffer;
  } catch (error) {
    console.error("Error generating and saving Excel file:", error);
  }
};
