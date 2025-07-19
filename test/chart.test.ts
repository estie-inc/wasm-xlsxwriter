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
  ChartSolidFill,
  ChartFormat,
  Color,
  ChartLine,
  ChartLayout,
  ChartGradientFill,
  ChartGradientStop,
  ChartGradientFillType,
  ChartPatternFill,
  ChartPatternFillType,
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
    const chartDataLabel1 = new ChartDataLabel().setFont(chartFont).showValue().setPosition(ChartDataLabelPosition.Left);

    const chartSeries1 = new ChartSeries();
    const chartLine1 = new ChartLine().setColor(Color.green());
    const chartSolidFill1 = new ChartSolidFill().setColor(Color.green());
    const chartFormat1 = new ChartFormat().setLine(chartLine1).setSolidFill(chartSolidFill1);
    const chartMarker1 = new ChartMarker().setType(ChartMarkerType.Circle).setSize(10).setFormat(chartFormat1);
    const categoriesRange1 = ChartRange.newFromRange("Sheet1", 1, 0, 6, 0);
    const valuesRange1 = ChartRange.newFromRange("Sheet1", 1, 1, 6, 1);
    chartSeries1
      .setName("Score 1")
      .setCategories(categoriesRange1)
      .setValues(valuesRange1)
      .setFormat(chartFormat1)
      .setDataLabel(chartDataLabel1)
      .setMarker(chartMarker1);

    const chartDataLabel2 = new ChartDataLabel();
    const chartSeries2 = new ChartSeries();
    const chartLine2 = new ChartLine();
    const chartSolidFill2 = new ChartSolidFill();
    const chartFormat2 = new ChartFormat();
    const chartMarker2 = new ChartMarker();
    const categoriesRange2 = ChartRange.newFromRange("Sheet1", 1, 0, 6, 0);
    const valuesRange2 = ChartRange.newFromRange("Sheet1", 1, 2, 6, 2);
    chartSeries2
      .setName("Score 2")
      .setCategories(categoriesRange2)
      .setValues(valuesRange2)
      .setFormat(chartFormat2.setLine(chartLine2.setColor(Color.purple())).setSolidFill(chartSolidFill2.setColor(Color.purple())))
      .setDataLabel(chartDataLabel2.setFont(chartFont).showValue().setPosition(ChartDataLabelPosition.Left))
      .setMarker(chartMarker2.setType(ChartMarkerType.Diamond).setSize(10).setFormat(chartFormat2.setLine(chartLine2.setColor(Color.purple())).setSolidFill(chartSolidFill2.setColor(Color.purple()))));
    chart
      .pushSeries(chartSeries1)
      .pushSeries(chartSeries2);
    const chartLayout = new ChartLayout().setOffset(0.1, 0.1);
    chart.title().setName("Score Transition").setFont(chartFont).setOverlay(true).setLayout(chartLayout);
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

  test("insert chart with gradient fill, pattern fill", async () => {
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

    const chart = new Chart(ChartType.Column);
    const categoriesRange = ChartRange.newFromRange("Sheet1", 1, 0, 6, 0);

    const chartSeries1 = new ChartSeries();
    const chartGradientFill1 = new ChartGradientFill().setType(ChartGradientFillType.Linear).setAngle(45).setGradientStops([new ChartGradientStop(Color.green(), 0), new ChartGradientStop(Color.yellow(), 100)]);
    const chartFormat1 = new ChartFormat().setGradientFill(chartGradientFill1);
    const valuesRange1 = ChartRange.newFromRange("Sheet1", 1, 1, 6, 1);
    chartSeries1
      .setName("Score 1")
      .setCategories(categoriesRange)
      .setValues(valuesRange1)
      .setFormat(chartFormat1);

    const chartSeries2 = new ChartSeries();
    const chartPatternFill = new ChartPatternFill().setPattern(ChartPatternFillType.DiagonalBrick).setBackgroundColor(Color.purple());
    const chartFormat2 = new ChartFormat().setPatternFill(chartPatternFill);
    const valuesRange2 = ChartRange.newFromRange("Sheet1", 1, 2, 6, 2);
    chartSeries2
      .setName("Score 2")
      .setCategories(categoriesRange)
      .setValues(valuesRange2)
      .setFormat(chartFormat2);
    chart
      .pushSeries(chartSeries1)
      .pushSeries(chartSeries2);

    worksheet.insertChart(0, 3, chart);

    // Assert
    const actual = await readXlsx(workbook.saveToBufferSync());
    const expected = await readXlsxFile("./expected/insert_chart_column.xlsx");
    expect(actual).matchXlsx(expected);
  });
});
