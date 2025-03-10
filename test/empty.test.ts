import { Workbook } from "..";
import { describe, test, expect } from "vitest";
import { readXlsx, readXlsxFile } from "./common";

describe("xlsx-wasm test", () => {
  test("create an empty workbook", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    // Do nothing

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/empty.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
