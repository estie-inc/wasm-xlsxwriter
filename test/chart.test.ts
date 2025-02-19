import {
  Workbook,
  Chart,
  ChartType,
  ChartSeries,
  ChartRange,
  ChartLegendPosition,
} from "../web/wasm_xlsxwriter";
import { describe, test, beforeAll, expect } from "vitest";
import { initWasModule, readXlsx, readXlsxFile } from "./common";

beforeAll(async () => {
  await initWasModule();
});

const DATA = [
  [2, 3, 4, 5, 6, 7],
  [10, 40, 50, 20, 10, 50],
  [30, 60, 70, 50, 40, 30],
];

describe("xlsx-wasm test", () => {
  test("insert chart", async () => {
    // Arrange
    const workbook = new Workbook();

    // Act
    const worksheet = workbook.addWorksheet();
    worksheet.write(0, 0, "Number");
    worksheet.write(0, 1, "Score 1");
    worksheet.write(0, 2, "Score 2");

    DATA.forEach((col_data, col_num) => {
      col_data.forEach((row_data, row_num) => {
        worksheet.write(row_num + 1, col_num, row_data);
      });
    });

    const chart = new Chart(ChartType.Line);

    const chartSeries1 = new ChartSeries();
    const categoriesRange1 = new ChartRange("Sheet1", 1, 0, 6, 0);
    const valuesRange1 = new ChartRange("Sheet1", 1, 1, 6, 1);
    chartSeries1
      .setName("Score 1")
      .setCategories(categoriesRange1)
      .setValues(valuesRange1);
    const chartSeries2 = new ChartSeries();
    const categoriesRange2 = new ChartRange("Sheet1", 1, 0, 6, 0);
    const valuesRange2 = new ChartRange("Sheet1", 1, 2, 6, 2);
    chartSeries2
      .setName("Score 2")
      .setCategories(categoriesRange2)
      .setValues(valuesRange2);
    chart
      .setName("Score Transition")
      .pushSeries(chartSeries1)
      .pushSeries(chartSeries2);
    chart.setWidth(640).setHeight(480);
    chart.xAxis().setName("x-axis");
    chart.yAxis().setName("y-axis");
    chart.legend().setPosition(ChartLegendPosition.Bottom);

    worksheet.insertChart(0, 3, chart);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/insert_chart.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
