import { Workbook, TableFunction, TableColumn, Formula, Table } from "..";
import { describe, test, expect } from "vitest";
import { readXlsx, readXlsxFile } from "./common";

describe("xlsx-wasm test", () => {
  test("create an empty workbook", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    const worksheet = workbook.addWorksheet();
    const items = ["Apples", "Pears", "Bananas", "Oranges"];
    const data = [
      [10000, 5000, 8000, 6000],
      [2000, 3000, 4000, 5000],
      [6000, 6000, 6500, 6000],
      [500, 300, 200, 700],
    ];
    // Write the table data.
    worksheet.writeColumn(3, 1, items);
    worksheet.writeRowMatrix(3, 2, data);
    // Create a new table and configure it.
    let columns = [
      new TableColumn().setHeader("Product").setTotalLabel("Totals"),
      new TableColumn()
        .setHeader("Quarter 1")
        .setTotalFunction(TableFunction.sum()),
      new TableColumn()
        .setHeader("Quarter 2")
        .setTotalFunction(TableFunction.sum()),
      new TableColumn()
        .setHeader("Quarter 3")
        .setTotalFunction(TableFunction.sum()),
      new TableColumn()
        .setHeader("Quarter 4")
        .setTotalFunction(TableFunction.sum()),
      new TableColumn()
        .setHeader("Year")
        .setTotalFunction(TableFunction.sum())
        .setFormula(new Formula("SUM(Table1[@[Quarter 1]:[Quarter 4]])")),
    ];
    const table = new Table().setColumns(columns).setTotalRow(true);
    worksheet.addTable(2, 1, 7, 6, table);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/table.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
