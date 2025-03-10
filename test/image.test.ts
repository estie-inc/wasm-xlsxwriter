import { Workbook, Image } from "..";
import { describe, test, expect } from "vitest";
import { loadFile, readXlsx, readXlsxFile } from "./common";

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

describe("xlsx-wasm test", () => {
  test("mutate image", async () => {
    // Arrange
    const workbook = new Workbook();
    const imageBuf = loadFile("./fixtures/rust.png");

    // Act
    const worksheet = workbook.addWorksheet();
    // method chaining
    const wideImage = new Image(imageBuf).setScaleWidth(2);
    // not method chaining
    const tallImage = new Image(imageBuf);
    tallImage.setScaleHeight(2);
    worksheet.insertImage(0, 0, wideImage);
    worksheet.insertImage(10, 0, tallImage);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/mutate_image.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
