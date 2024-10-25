import { Workbook, Image } from "../rust/pkg";
import { describe, test, beforeAll, expect } from "vitest";
import { initWasModule, loadFile, readXlsx, readXlsxFile } from "./common";

beforeAll(async () => {
  await initWasModule();
});

describe("xlsx-wasm test", () => {
  test("insert image", async () => {
    // Arrange
    const workbook = new Workbook();
    const imageBuf = loadFile("./fixtures/rust.png");
    const image = new Image(imageBuf);

    // Act
    const worksheet = workbook.addWorksheet();
    worksheet.insertImage(0, 0, image);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/insert_image.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test("embed image", async () => {
    // Arrange
    const workbook = new Workbook();
    const imageBuf = loadFile("./fixtures/rust.png");
    const image = new Image(imageBuf);

    // Act
    const worksheet = workbook.addWorksheet();
    worksheet.embedImage(1, 1, image);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/embed_image.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
