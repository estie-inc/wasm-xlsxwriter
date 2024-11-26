# wasm-xlsxwriter [![NPM Version](https://img.shields.io/npm/v/wasm-xlsxwriter)](https://www.npmjs.com/package/wasm-xlsxwriter)

The `wasm-xlsxwriter` library is a lightweight wrapper around the write API of Rust's [`rust_xlsxwriter`](https://crates.io/crates/rust_xlsxwriter), compiled to WebAssembly (Wasm) with minimal setup to make it easily usable from JavaScript or Node.js.

With this library, you can generate Excel files in the browser or Node.js using JavaScript, complete with support for custom formatting, formulas, links, images, and more.

## Getting Started

To get started, install the library via npm:

```bash
npm install wasm-xlsxwriter
```

### Usage Web

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
const image = new Image(imageBuffer);
worksheet.insertImage(1, 2, image);

// Save the file to a buffer.
const buf = workbook.saveToBufferSync();
```

### Usage Node.js
```js
const {
	Format,
	FormatAlign,
	FormatBorder,
	FormatScript,
	DocProperties,
	Color,
	Workbook
} = require('wasm-xlsxwriter');

/**
 * Demo function that create a xlsx buffer from a header array and data rows 
 *
 * @param {Array.<String>} header - Sheet header strings
 * @param {Array.<Array<*>>} rows - Array of arrays containing sheet rows
 * @returns {Buffer} xlsx file data
 */
function exportXlsx(header, rows) {
	// Create a new Excel file object.
	const workbook = new Workbook();

	// Set a doc property
	const props = new DocProperties();
	props.setCompany('Test Inc.');
	workbook.setProperties(props);

	// Add a worksheet with name to the workbook.
	const worksheet = workbook.addWorksheet();
	worksheet.setName('Export');

	// Create some formats to use in the worksheet.
	const headerStyle = new Format();
	headerStyle.setAlign(FormatAlign.Top).setAlign(FormatAlign.Left).setTextWrap().setBackgroundColor(Color.rgb(0xe2eee8));
	headerStyle.setBorderBottom(FormatBorder.Medium).setBorderBottomColor(Color.rgb(0x0B703D));
	headerStyle.setBorderRight(FormatBorder.Thin).setBorderRightColor(Color.rgb(0x0B703D));
	const defaultStyle = new Format().setTextWrap().setAlign(FormatAlign.Top);

	// Write sheet header
	for (const [x, col] of header.entries()) {
		worksheet.writeWithFormat(0, x, col, headerStyle);
	}

	// Write sheet data
	for (const [y, row] of rows.entries()) {
		for (const [x, col] of row.entries()) {
			worksheet.writeWithFormat(y + 1, x, col, defaultStyle);
		}
	}

	// Autofit columns
	worksheet.autofit();

	// Freeze header
	worksheet.setFreezePanes(1, 0);

	// Add autofilter to header
	worksheet.autofilter(0, 0, rows.length, header.length - 1);

	// Return buffer with xlsx data
	return workbook.saveToBufferSync();
});
```


## License

MIT
