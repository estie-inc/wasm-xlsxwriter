import { Workbook } from "../rust/pkg";
import { describe, test, beforeAll, expect } from "vitest";
import { initWasModule, readXlsx, readXlsxFile } from "./common";

beforeAll(async () => {
  await initWasModule();
});

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
