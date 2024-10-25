import "vitest";
import { XlsxFile } from "./test/common";

interface CustomMatchers<R = unknown> {
  matchXlsx: (xlsxFile: XlsxFile) => R;
}

declare module "vitest" {
  interface Assertion<T = any> extends CustomMatchers<T> {}
  interface AsymmetricMatchersContaining extends CustomMatchers {}
}
