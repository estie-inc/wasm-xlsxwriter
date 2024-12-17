import {
  Color,
  DocProperties,
  Format,
  FormatAlign,
  Workbook,
} from "wasm-xlsxwriter";
import * as fs from "fs";

/**
 * Demo function that create a xlsx buffer from a header array and data rows
 *
 * @param {string[]} header - Sheet header strings
 * @param {(string | number)[][]} rows - Array of arrays containing sheet rows
 * @returns {Buffer} Excel file data
 */
function writeExcel(header: string[], rows: (string | number)[][]): Buffer {
  // Create a new Excel file object.
  const workbook = new Workbook();

  // Set a doc property
  const props = new DocProperties().setCompany("Test, Inc.");
  workbook.setProperties(props);

  // Add a worksheet with name to the workbook.
  const worksheet = workbook.addWorksheet();
  worksheet.setName("Export");

  // Create some formats to use in the worksheet.
  const headerStyle = new Format();
  headerStyle
    .setAlign(FormatAlign.Top)
    .setTextWrap()
    .setBackgroundColor(Color.red());

  // Write sheet header
  worksheet.writeRowWithFormat(0, 0, header, headerStyle);

  // Write sheet data
  worksheet.writeRowMatrix(1, 0, rows);

  // Autofit columns
  worksheet.autofit();

  // Freeze header
  worksheet.setFreezePanes(1, 0);

  // Add autofilter to header
  worksheet.autofilter(0, 0, rows.length, header.length - 1);

  // Return buffer with xlsx data
  const uint8Array = workbook.saveToBufferSync();
  return Buffer.from(uint8Array);
}

const buf = writeExcel(
  ["Name", "Age", "City"],
  [
    ["John", 30, "New York"],
    ["Jane", 25, "Los Angeles"],
  ]
);
fs.writeFileSync("export.xlsx", buf);
