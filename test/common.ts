import fs from "fs";
import path from "path";
import { assert } from "console";
import unzipper from "unzipper";
import initWasmBindgen from "../web";

export async function initWasModule() {
  const wasmSource = await fs.promises.readFile("web/wasm_xlsxwriter_bg.wasm");
  const wasmModule = await WebAssembly.compile(wasmSource);
  await initWasmBindgen({ module_or_path: wasmModule });
}

export function loadFile(relativePath: string): Buffer {
  return fs.readFileSync(path.resolve(__dirname, relativePath));
}

export function saveFile(relativePath: string, data: Uint8Array): void {
  fs.writeFileSync(path.resolve(__dirname, relativePath), data);
}

export const readXlsxFile = async (path: string) => {
  const buf = loadFile(path);
  return readXlsx(buf);
};

export const readXlsx = async (buf: Uint8Array) => {
  const files = new Map<string, string>();

  const zipEntries = await unzip(Buffer.from(buf));
  if (zipEntries.length === 0) {
    throw new Error("No zip entries");
  }

  for (const zipEntry of zipEntries) {
    let data = zipEntry.data.toString();

    // https://github.com/jmcnamara/rust_xlsxwriter/blob/f1c1c1f3e7809f1422260e269adb25022531f81e/tests/integration/common/mod.rs#L315
    if (zipEntry.entryName === "docProps/core.xml") {
      // Remove creation timestamp
      data = data.replace(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z/g, "");
    } else if (zipEntry.entryName === "xl/workbook.xml") {
      // Remove workbookView dimensions which are almost always different and
      // calcPr which can have different Excel version ids.
      data = data.replace(
        /<workbookView xWindow="\d+" yWindow="\d+" windowWidth="\d+" windowHeight="\d+"/g,
        "<workbookView"
      );
    } else if (zipEntry.entryName.startsWith("xl/charts/chart")) {
      // The pageMargins element in chart files often contain values like
      // "0.75000000000000011" instead of "0.75". We simplify/round these to
      // make comparison easier.
      data = data.replace(/000000000000\d+/g, "");
    }

    // TODO: do somehing if rust_xlsxwriter is built with feature = "ryu"
    // https://github.com/jmcnamara/rust_xlsxwriter/blob/f1c1c1f3e7809f1422260e269adb25022531f81e/tests/integration/common/mod.rs#L356

    files.set(zipEntry.entryName, data);
  }

  return { files };
};

export interface XlsxFile {
  files: Map<string, string>;
}

export function compareXlsxFiles(a: XlsxFile, b: XlsxFile): boolean {
  if (a.files.size !== b.files.size) {
    return false;
  }

  for (const [key, value] of a.files) {
    if (!b.files.has(key)) {
      return false;
    }

    if (value !== b.files.get(key)) {
      return false;
    }
  }

  return true;
}

async function unzip(buf: Buffer): Promise<
  {
    entryName: string;
    data: Buffer;
  }[]
> {
  assert(
    buf[0] && buf[1] && String.fromCodePoint(buf[0], buf[1]) === "PK",
    "Not a zip file"
  );
  const dirs = await unzipper.Open.buffer(buf);
  const ret = [];
  for (const f of dirs.files) {
    const data = await f.buffer();
    ret.push({ entryName: f.path, data });
  }
  return ret;
}
