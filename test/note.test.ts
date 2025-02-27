import { Color, Format, Note, Workbook } from "../web";
import { describe, test, beforeAll, expect } from "vitest";
import { initWasModule, readXlsx, readXlsxFile } from "./common";

beforeAll(async () => {
  await initWasModule();
});

describe("xlsx-wasm test", () => {
  test("insert note", async () => {
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
    worksheet.insertNote(0, 0, new Note('This is a note'));

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/insert_note.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
