# wasm-xlsxwriter [![NPM Version](https://img.shields.io/npm/v/wasm-xlsxwriter)](https://www.npmjs.com/package/wasm-xlsxwriter)

The `wasm-xlsxwriter` library is a lightweight wrapper around the write API of Rust's [`rust_xlsxwriter`](https://crates.io/crates/rust_xlsxwriter), compiled to WebAssembly (Wasm) with minimal setup to make it easily usable from JavaScript.

With this library, you can generate Excel files in the browser using JavaScript, complete with support for custom formatting, formulas, links, images, and more.

## Getting Started

To get started, install the library via npm:

```bash
npm install wasm-xlsxwriter
```

### Usage

Here’s an example of how to use `wasm-xlsxwriter` to create an Excel file:

```typescript
import xlsxInit, {
  Format,
  FormatAlign,
  FormatBorder,
  Formula,
  Workbook,
  Image,
  Url,
} from "wasm-xlsxwriter";

// Load the WebAssembly module and initialize the library.
await xlsxInit();

// Create a new Excel file object.
const workbook = new Workbook();

// Create some formats to use in the worksheet.
const boldFormat = new Format().setBold();
const decimalFormat = new Format().setNumFormat("0.000");
const dateFormat = new Format().setNumFormat("yyyy-mm-dd");
const mergeFormat = new Format()
  .setBorder(FormatBorder.Thin)
  .setAlign(FormatAlign.Center);

// Add a worksheet to the workbook.
const worksheet = workbook.addWorksheet();

// Set the column width for clarity.
worksheet.setColumnWidth(0, 22);

// Write a string without formatting.
worksheet.write(0, 0, "Hello");

// Write a string with the bold format defined above.
worksheet.writeWithFormat(1, 0, "World", boldFormat);

// Write some numbers.
worksheet.write(2, 0, 1);
worksheet.write(3, 0, 2.34);

// Write a number with formatting.
worksheet.writeWithFormat(4, 0, 3.0, decimalFormat);

// Write a formula.
worksheet.write(5, 0, new Formula("=SIN(PI()/4)"));

// Write a date.
const date = new Date(2023, 1, 25);
worksheet.writeWithFormat(6, 0, date, dateFormat);

// Write some links.
worksheet.write(7, 0, new Url("https://www.rust-lang.org"));
worksheet.write(8, 0, new Url("https://www.rust-lang.org").setText("Rust"));

// Write some merged cells.
worksheet.mergeRange(9, 0, 9, 1, "Merged cells", mergeFormat);

// Insert an image (ensure `imageBuffer` contains the image data).
const image = Image.newFromBuffer(imageBuffer);
worksheet.insertImage(1, 2, image);

// Save the file to a buffer.
const buf = workbook.saveToBufferSync();
```

## License

MIT
