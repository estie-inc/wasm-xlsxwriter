import {
  Workbook,
  Chart,
  ChartType,
  ChartSeries,
  ChartRange,
  ChartLegendPosition,
  ChartFont,
  ChartDataLabel,
  ChartDataLabelPosition,
  ChartMarker,
  ChartMarkerType,
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
  
    const chart = new Chart(ChartType.Stock);
    const chartFont = new ChartFont().setName("Meiryo UI");
    const chartDataLabel = new ChartDataLabel().setFont(chartFont).showValue().setPosition(ChartDataLabelPosition.Left);
  
    const chartSeries1 = new ChartSeries();
    const chartMarker1 = new ChartMarker().setType(ChartMarkerType.Circle).setSize(10);
    const categoriesRange1 = new ChartRange("Sheet1", 1, 0, 6, 0);
    const valuesRange1 = new ChartRange("Sheet1", 1, 1, 6, 1);
    chartSeries1
      .setName("Score 1")
      .setCategories(categoriesRange1)
      .setValues(valuesRange1)
      .setDataLabel(chartDataLabel)
      .setMarker(chartMarker1);
    const chartSeries2 = new ChartSeries();
    const chartMarker2 = new ChartMarker().setType(ChartMarkerType.Diamond).setSize(10);
    const categoriesRange2 = new ChartRange("Sheet1", 1, 0, 6, 0);
    const valuesRange2 = new ChartRange("Sheet1", 1, 2, 6, 2);
    chartSeries2
      .setName("Score 2")
      .setCategories(categoriesRange2)
      .setValues(valuesRange2)
      .setDataLabel(chartDataLabel)
      .setMarker(chartMarker2);
    chart
      .pushSeries(chartSeries1)
      .pushSeries(chartSeries2);
    chart.title().setName("Score Transition").setFont(chartFont);
    chart.setWidth(640).setHeight(480);
    chart.xAxis().setName("x-axis").setFont(chartFont).setNameFont(chartFont);
    chart.yAxis().setName("y-axis").setFont(chartFont).setNameFont(chartFont);
    chart.legend().setPosition(ChartLegendPosition.Bottom).setFont(chartFont);

    worksheet.insertChart(0, 3, chart);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/insert_chart.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
