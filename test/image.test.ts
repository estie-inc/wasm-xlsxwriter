import { Workbook, Image, ObjectMovement } from "../web";
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
    image.setAltText("rust logo");
    image.setObjectMovement(ObjectMovement.MoveAndSizeWithCells);
    worksheet.insertImage(0, 0, image);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/insert_image.xlsx");
    expect(actual).matchXlsx(expected);
  });
});

describe("xlsx-wasm test", () => {
  test("insert image centered", async () => {
    // Arrange
    const workbook = new Workbook();
    const imageBuf = loadFile("./fixtures/rust.png");
    const image = new Image(imageBuf);

    // Act
    const worksheet = workbook.addWorksheet();
    worksheet.setColumnWidthPixels(0, 200);
    worksheet.setRowHeight(0, 100);
    worksheet.insertImageFitToCellCentered(0, 0, image);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/insert_image_fit_centered.xlsx");
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
    image.setDecorative(true);
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
    // set width, height
    const image = new Image(imageBuf);
    image.setWidth(50);
    image.setHeight(50);
    worksheet.insertImage(20, 0, image);
    // set scale to size
    const scaleImage = new Image(imageBuf);
    scaleImage.setScaleToSize(50, 100, false);
    worksheet.insertImage(30, 0, scaleImage);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/mutate_image.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
