import { Workbook } from "wasm-xlsxwriter/web";
import { describe, test, beforeAll, expect } from "vitest";
import { initWasModule, readXlsx, readXlsxFile } from "./common";

beforeAll(async () => {
  await initWasModule();
});

describe("xlsx-wasm test", () => {
  test("group rows with symbols above", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    // Heading rows 0 and 4, detail rows 1-3 and 5-7.
    const worksheet = workbook.addWorksheet();
    worksheet.write(0, 0, "Group 1");
    worksheet.write(1, 0, "Item 1-1");
    worksheet.write(1, 1, 100);
    worksheet.write(2, 0, "Item 1-2");
    worksheet.write(2, 1, 200);
    worksheet.write(3, 0, "Item 1-3");
    worksheet.write(3, 1, 300);
    worksheet.write(4, 0, "Group 2");
    worksheet.write(5, 0, "Item 2-1");
    worksheet.write(5, 1, 400);
    worksheet.write(6, 0, "Item 2-2");
    worksheet.write(6, 1, 500);
    worksheet.write(7, 0, "Item 2-3");
    worksheet.write(7, 1, 600);
    worksheet.groupRows(1, 3);
    worksheet.groupRows(5, 7);
    worksheet.groupSymbolsAbove(true);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    expect(actual.files.get("xl/worksheets/sheet1.xml")).toContain(
      '<sheetPr><outlinePr summaryBelow="0"/></sheetPr>'
    );
    const expected = await readXlsxFile("./expected/outline_symbols_above.xlsx");
    expect(actual).matchXlsx(expected);
  });

  test("group rows collapsed", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    // Detail rows 1-3 and a total row 4.
    const worksheet = workbook.addWorksheet();
    worksheet.write(1, 0, "Item 1");
    worksheet.write(1, 1, 100);
    worksheet.write(2, 0, "Item 2");
    worksheet.write(2, 1, 200);
    worksheet.write(3, 0, "Item 3");
    worksheet.write(3, 1, 300);
    worksheet.write(4, 0, "Total");
    worksheet.write(4, 1, 600);
    worksheet.groupRowsCollapsed(1, 3);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const sheet = actual.files.get("xl/worksheets/sheet1.xml");
    // The grouped rows are hidden and the row after the group is collapsed.
    expect(sheet).toMatch(/<row r="2" [^>]*hidden="1" outlineLevel="1">/);
    expect(sheet).toMatch(/<row r="4" [^>]*hidden="1" outlineLevel="1">/);
    expect(sheet).toMatch(/<row r="5" [^>]*collapsed="1">/);
    const expected = await readXlsxFile("./expected/outline_rows_collapsed.xlsx");
    expect(actual).matchXlsx(expected);
  });

  test("group rows collapsed with symbols above", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    // Heading rows 0 and 4, detail rows 1-3 and 5-7.
    const worksheet = workbook.addWorksheet();
    worksheet.write(0, 0, "Group 1");
    worksheet.write(1, 0, "Item 1-1");
    worksheet.write(2, 0, "Item 1-2");
    worksheet.write(3, 0, "Item 1-3");
    worksheet.write(4, 0, "Group 2");
    worksheet.write(5, 0, "Item 2-1");
    worksheet.write(6, 0, "Item 2-2");
    worksheet.write(7, 0, "Item 2-3");
    worksheet.groupRowsCollapsed(1, 3);
    worksheet.groupRowsCollapsed(5, 7);
    worksheet.groupSymbolsAbove(true);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const sheet = actual.files.get("xl/worksheets/sheet1.xml");
    // rust_xlsxwriter writes collapsed on last_row + 1 (r="5" and r="9")
    // even when the symbols are above, while Excel writes it on the heading
    // rows (r="1" and r="5"). If this fails after a rust_xlsxwriter update,
    // revisit the note in the groupRowsCollapsed doc comment.
    expect(sheet).not.toMatch(/<row r="1" [^>]*collapsed="1"/);
    expect(sheet).toMatch(/<row r="5" [^>]*collapsed="1">/);
    expect(sheet).toMatch(/<row r="9" [^>]*collapsed="1"\/>/);
  });
});
