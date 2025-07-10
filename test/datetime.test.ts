import { ExcelDateTime } from "../web";
import { describe, test, expect, beforeAll } from "vitest";
import { initWasModule } from "./common.js";

beforeAll(async () => {
  await initWasModule();
});

describe("ExcelDateTime Normal Cases", () => {
  test.each([
    [2024, 6, 1, 45444],
    [2024, 2, 29, 45351],
    [2024, 1, 1, 45292],
    [2024, 12, 31, 45657],
    [9999, 12, 31, 2958465],
    [1899, 12, 31, 0],
  ])("fromYMD(%i, %i, %i) should create instance and toString() === %s", (y, m, d, expectedStr) => {
    const dt = ExcelDateTime.fromYMD(y, m, d);
    expect(dt).toBeInstanceOf(ExcelDateTime);
    expect(dt.toExcel()).toBe(expectedStr);
  });

  test.each([
    ["2024-06-01", 45444],
    ["2024-02-29", 45351],
    ["2024-01-01", 45292],
    ["2024-12-31", 45657],
    ["9999-12-31", 2958465],
    ["1899-12-31", 0],
  ])("parseFromStr('%s') should create instance and toString() === %s", (str, expectedStr) => {
    const dt = ExcelDateTime.parseFromStr(str);
    expect(dt).toBeInstanceOf(ExcelDateTime);
    expect(dt.toExcel()).toBe(expectedStr);
  });
});

describe("ExcelDateTime Error Cases", () => {
  test.each([
    [2024, 2, 30],
    [1899, 12, 30],
  ])("fromYMD(%i, %i, %i) should throw", (y, m, d) => {
    expect(() => ExcelDateTime.fromYMD(y, m, d)).toThrow(/DateTimeRangeError/);
  });

  test.each([
    ["2022"],
    ["2024/06/01"], // Different delimiter
    ["2024.06.01"], // Different delimiter
  ])("parseFromStr('%s') should throw ParseError", (str) => {
    expect(() => ExcelDateTime.parseFromStr(str)).toThrow(/DateTimeParseError/);
  });

  test.each([
    ["not-a-date"],
    ["2024-02-30"],
    ["1899-12-30"],
  ])("parseFromStr('%s') should throw RangeError", (str) => {
    expect(() => ExcelDateTime.parseFromStr(str)).toThrow(/DateTimeRangeError/);
  });
}); 