import { Color, Format, Note, Workbook, ObjectMovement } from "../web";
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
    const note1 = new Note("This is a note")
      .setAuthor("Yuya Ryuzaki")
      .addAuthorPrefix(false)
      .setWidth(200)
      .setHeight(100)
      .setBackgroundColor(Color.purple())
      .setFontName("Meiryo UI")
      .setFontSize(20)
      .setAltText("Alt text")
      .setObjectMovement(ObjectMovement.DontMoveOrSizeWithCells);
    note1.resetText("This is other note");
    const format = new Format()
      .setBackgroundColor(Color.green())
      .setFontName("Meiryo UI")
      .setFontSize(11);
    const note2 = new Note("This is a note")
      .addAuthorPrefix(true)
      .setWidth(100)
      .setHeight(100)
      .setFormat(format)
      .setVisible(true)
      .setObjectMovement(ObjectMovement.MoveAndSizeWithCells);

    worksheet.insertNote(0, 0, note1);
    worksheet.insertNote(0, 1, note2);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/insert_note.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
