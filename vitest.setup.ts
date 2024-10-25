import { compareXlsxFiles, XlsxFile } from "./test/common";
import { expect } from "vitest";

expect.extend({
  matchXlsx(
    received: XlsxFile,
    expected: XlsxFile
  ): { message: () => string; pass: boolean } {
    const pass = compareXlsxFiles(received, expected);
    if (pass) {
      return {
        message: () => `expected ${received} not to have value ${expected}`,
        pass: true,
      };
    } else {
      return {
        message: () => `expected ${received} to have value ${expected}`,
        pass: false,
      };
    }
  },
});
