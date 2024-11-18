import {
  Color,
  Format,
  FormatAlign,
  FormatBorder,
  FormatUnderline,
  Formula,
  Workbook,
} from "../rust/pkg";
import { describe, test, beforeAll, expect } from "vitest";
import { initWasModule, readXlsx, readXlsxFile } from "./common";

beforeAll(async () => {
  await initWasModule();
});

describe("xlsx-wasm test", () => {
  test("use format", async () => {
    // Arrange
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    // Act
    worksheet.writeStringWithFormat(0, 0, "bold", new Format().setBold());

    const format1 = new Format().setItalic();
    worksheet.writeStringWithFormat(0, 1, "italic", format1);
    worksheet.writeNumberWithFormat(0, 2, 123, format1);
    worksheet.writeBooleanWithFormat(0, 3, true, format1);

    const format2 = new Format().setRotation(45).setAlign(FormatAlign.Center);
    worksheet.writeStringWithFormat(0, 4, "rot45", format2);

    const format3 = new Format().setFontStrikethrough();
    worksheet.writeStringWithFormat(0, 5, "strikethrough", format3);

    const format4 = new Format().setUnderline(FormatUnderline.Single);
    worksheet.writeStringWithFormat(0, 6, "underline single", format4);

    const format5 = new Format().setIndent(2);
    worksheet.writeStringWithFormat(0, 7, "indent2", format5);

    const format6 = new Format().setHidden();
    worksheet.writeFormulaWithFormat(0, 8, new Formula("=PI()"), format6);
    worksheet.protect();

    const format7 = new Format().setFontSize(20);
    worksheet.writeStringWithFormat(0, 9, "font size 20", format7);

    const format8 = new Format().setFontFamily(2);
    worksheet.writeStringWithFormat(0, 10, "font family 2", format8);

    const format9 = new Format().setFontName("Arial");
    worksheet.writeStringWithFormat(0, 11, "font name Arial", format9);

    const format10 = new Format().setFontCharset(1);
    worksheet.writeStringWithFormat(0, 12, "font charset 1", format10);

    const numFormat = new Format().setNumFormat("yyyy/m/d h:mm");
    worksheet.writeDatetimeWithFormat(
      0,
      4,
      new Date(Date.UTC(2000, 12, 12)),
      numFormat
    );
    worksheet.writeDatetimeWithFormat(
      0,
      5,
      new Date(Date.UTC(2030, 8, 30, 23, 59, 59)),
      numFormat
    );

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/format.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test("use format border", async () => {
    // Arrange
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    // Act
    const format1 = new Format().setBorder(FormatBorder.Dotted);
    worksheet.writeStringWithFormat(0, 0, "AAA", format1);

    const format2 = new Format()
      .setBorderColor(Color.rgb(0x00ff00))
      .setBorderBottom(FormatBorder.Dashed);
    worksheet.writeStringWithFormat(0, 1, "BBB", format2);

    const format3 = new Format()
      .setBorderTop(FormatBorder.Dotted)
      .setBorderRight(FormatBorder.Dashed);
    worksheet.writeStringWithFormat(0, 2, "CCC", format3);

    const format4 = new Format()
      .setBorderDiagonal(FormatBorder.Dotted)
      .setBorderDiagonalColor(Color.rgb(0xff0000));
    worksheet.writeStringWithFormat(0, 3, "DDD", format4);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/format_border.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test("use color", async () => {
    // Arrange
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    // Act
    const black = Color.black();
    const purple = Color.purple();
    const blue = Color.rgb(0x0000ff);
    const red = Color.parse("#FF0000");
    const yellow = Color.yellow();

    worksheet.writeStringWithFormat(
      0,
      0,
      "A",
      new Format().setFontColor(black)
    );
    worksheet.writeStringWithFormat(
      0,
      1,
      "B",
      new Format().setFontColor(purple)
    );
    worksheet.writeStringWithFormat(
      0,
      2,
      "C",
      new Format().setBackgroundColor(blue)
    );
    worksheet.writeStringWithFormat(
      0,
      3,
      "D",
      new Format().setBackgroundColor(red)
    );
    worksheet.writeStringWithFormat(
      0,
      4,
      "E",
      new Format().setForegroundColor(yellow)
    );

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/format_color.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test('use width and height', async () => {
    // Arrange
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    // Act
    worksheet.setColumnWidth(0, 16);
    worksheet.setColumnRangeWidth(1, 2, 32);
    worksheet.setColumnWidthPixels(4, 100);

    worksheet.setRowHeight(0, 24);
    worksheet.setRowHeightPixels(1, 100); // 100 pixels = 75

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/format_width_height.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test("methods mutate self", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    const worksheet = workbook.addWorksheet();
    const format = new Format();
    format
      .setFontName('Meiryo UI')
      .setFontSize(16)
      .setAlign(FormatAlign.Center)
      .setBorder(FormatBorder.Thin);

    worksheet.writeWithFormat(0, 0, 'foo', format);
    worksheet.writeWithFormat(1, 1, 'bar', format);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/format_mutate.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test("clone format object", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    const worksheet = workbook.addWorksheet();
    const baseFormat = new Format().setBold();
    const format1 = baseFormat.clone().setItalic();
    const format2 = baseFormat.clone().setFontColor(Color.red());

    worksheet.writeStringWithFormat(0, 0, "bold", baseFormat);
    worksheet.writeStringWithFormat(0, 1, "bold italic", format1);
    worksheet.writeStringWithFormat(0, 2, "bold red", format2);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/format_clone.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
