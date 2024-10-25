import { Workbook, Format, Formula } from "../rust/pkg";
import { describe, test, beforeAll, expect } from "vitest";
import { initWasModule, readXlsx, readXlsxFile } from "./common";

beforeAll(async () => {
  await initWasModule();
});

describe("xlsx-wasm test", () => {
  test("write primitive", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    const worksheet = workbook.addWorksheet();
    worksheet.write(0, 0, "日本語");
    worksheet.write(1, 0, true);

    worksheet.write(2, 0, 3.14159265);
    worksheet.write(3, 0, 987654321);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/write_primitive.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test("write formula", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    const worksheet = workbook.addWorksheet();
    worksheet.write(0, 0, new Formula("=PI()"));

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/write_formula.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test("write date", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    const worksheet = workbook.addWorksheet();
    const format = new Format().setNumFormat("yyyy/m/d h:mm");
    worksheet.writeWithFormat(
      0,
      0,
      new Date(Date.UTC(1904, 1, 1, 12, 34, 56)),
      format
    );
    worksheet.writeWithFormat(
      1,
      0,
      new Date(Date.UTC(2030, 8, 30, 23, 59, 59)),
      format
    );

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/write_date.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test("write columns and rows", async () => {
    // Arrange
    const workbook = new Workbook();
    const items = [true, "some text", 10000];

    // Act
    const worksheet = workbook.addWorksheet();
    worksheet.writeRow(0, 0, items);
    worksheet.writeColumn(1, 0, items);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/write_column.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test("write column and row matrix", async () => {
    // Arrange
    const workbook = new Workbook();
    const matrix = [
      [true, "text1", 10000],
      [false, "text2", 20000],
    ];

    // Act
    const worksheet = workbook.addWorksheet();
    // write matrix as is
    worksheet.writeRowMatrix(0, 0, matrix);
    // write transposed matrix
    worksheet.writeColumnMatrix(4, 4, matrix);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/write_matrix.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
