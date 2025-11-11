import { beforeAll, describe, expect, test } from "vitest";
import { Format, Image, Workbook } from "../web";
import { initWasModule, loadFile, readXlsx, readXlsxFile } from "./common";

beforeAll(async () => {
  await initWasModule();
});


describe("xlsx-wasm default format test", () => {
  test("set default format", async () => {
    // Arrange
    const workbook = new Workbook();
    const format = new Format().setFontName('ＭＳ Ｐゴシック').setFontSize(11);

    const imageBuf = loadFile("./fixtures/rust.png");
    const image = new Image(imageBuf);

    // Act
    workbook.setDefaultFormat(format, 18, 72);

    const worksheet = workbook.addWorksheet();
    worksheet.write(0, 0, 'Test with default format');
    worksheet.write(1, 0, 123);

    worksheet.setColumnWidthPixels(1, 200);
    worksheet.setRowHeight(1, 100);
    worksheet.insertImageFitToCellCentered(1, 1, image);

    worksheet.setColumnWidthPixels(2, 100);
    worksheet.setRowHeight(2, 200);
    worksheet.insertImageFitToCellCentered(2, 2, image);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/default_format.xlsx");
    expect(actual).matchXlsx(expected);
  });

  test("set default format error - after worksheet added", async () => {
    // Arrange
    const workbook = new Workbook();
    const format = new Format().setFontName('Arial').setFontSize(10);

    // Act
    workbook.addWorksheet();

    // Assert
    expect(() => {
      workbook.setDefaultFormat(format, 15, 64);
    }).toThrow("XlsxError(DefaultFormatError(\"Default format must be set before adding worksheets.\"))");
  });

  test("set default format error - invalid column width", async () => {
    // Arrange
    const workbook = new Workbook();
    const format = new Format().setFontName('Arial').setFontSize(10);

    // Assert
    expect(() => {
      workbook.setDefaultFormat(format, 15, 999);
    }).toThrow("XlsxError(DefaultFormatError(\"Unsupported default column width: 999\"))");
  });
});