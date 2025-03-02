import { Color, Format, Workbook } from "wasm-xlsxwriter";
import { describe, test, expect } from "vitest";
import { readXlsx, readXlsxFile } from "./common";

describe("xlsx-wasm test", () => {
  test("worksheet print", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    const worksheet = workbook.addWorksheet();

    const purple = Color.purple();
    const format = new Format().setFontColor(purple);
    worksheet.writeWithFormat(0, 0, "Hello, World!", format);

    worksheet.setLandscape();
    worksheet.setScreenGridlines(false);
    worksheet.setPrintGridlines(true);
    worksheet.setPrintArea(0, 0, 10, 10);
    worksheet.setPrintScale(110);
    worksheet.setPaperSize(9); // A4
    worksheet.setPrintBlackAndWhite(true);
    worksheet.setPrintFirstPageNumber(0);
    worksheet.setPrintDraft(true);
    worksheet.setPrintCenterHorizontally(true);
    worksheet.setPrintCenterVertically(true);
    worksheet.setPrintHeadings(true);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/print.xlsx");
    expect(actual).matchXlsx(expected);
  });

  test("worksheet print portrait", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    const worksheet = workbook.addWorksheet();
    worksheet.write(0, 0, "Hello, World!");
    worksheet.setPortrait();

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/print_portrait.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
