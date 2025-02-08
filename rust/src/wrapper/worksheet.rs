use std::sync::{Arc, Mutex};

use crate::error::XlsxError;
use rust_xlsxwriter as xlsx;
use wasm_bindgen::{prelude::*, JsValue};

use crate::wrapper::{
    chart::Chart, excel_data::ExcelData, format::Format, image::Image, map_xlsx_error,
    table::Table, utils, WasmResult,
};

use super::{
    excel_data::{JsExcelData, JsExcelDataArray, JsExcelDataMatrix},
    formula::Formula,
    rich_string::RichString,
    url::Url,
};

/// The `Worksheet` struct represents an Excel worksheet. It handles operations
/// such as writing data to cells or formatting the worksheet layout.
///
/// TODO: example omitted
#[wasm_bindgen]
pub struct Worksheet {
    pub(crate) workbook: Arc<Mutex<xlsx::Workbook>>,
    pub(crate) index: usize,
}

impl Clone for Worksheet {
    fn clone(&self) -> Self {
        Worksheet {
            workbook: Arc::clone(&self.workbook),
            index: self.index,
        }
    }
}

#[wasm_bindgen]
impl Worksheet {
    /// Get the worksheet name.
    ///
    /// Get the worksheet name that was set automatically such as Sheet1,
    /// Sheet2, etc., or that was set by the user using
    /// {@link Worksheet#setName}.
    ///
    /// The worksheet name can be used to get a reference to a worksheet object
    /// using the {@link Workbook#worksheetFromName} method.
    ///
    /// TODO: example omitted
    #[wasm_bindgen]
    pub fn name(&self) -> String {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        sheet.name().to_string()
    }

    /// Set the worksheet name.
    ///
    /// Set the worksheet name. If no name is set the default Excel convention
    /// will be followed (Sheet1, Sheet2, etc.) in the order the worksheets are
    /// created.
    ///
    /// @param {string} name - The worksheet name. It must follow the Excel rules, shown
    ///   below.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::SheetnameCannotBeBlank`] - Worksheet name cannot be
    ///   blank.
    /// - [`XlsxError::SheetnameLengthExceeded`] - Worksheet name exceeds
    ///   Excel's limit of 31 characters.
    /// - [`XlsxError::SheetnameContainsInvalidCharacter`] - Worksheet name
    ///   cannot contain invalid characters: `[ ] : * ? / \`
    /// - [`XlsxError::SheetnameStartsOrEndsWithApostrophe`] - Worksheet name
    ///   cannot start or end with an apostrophe.
    ///
    /// TODO: example omitted
    ///
    /// The worksheet name must be a valid Excel worksheet name, i.e:
    ///
    /// - The name is less than 32 characters.
    /// - The name isn't blank.
    /// - The name doesn't contain any of the characters: `[ ] : * ? / \`.
    /// - The name doesn't start or end with an apostrophe.
    /// - The name shouldn't be "History" (case-insensitive) since that is
    ///   reserved by Excel.
    /// - It must not be a duplicate of another worksheet name used in the
    ///   workbook.
    ///
    /// The rules for worksheet names in Excel are explained in the [Microsoft
    /// Office documentation].
    ///
    /// [Microsoft Office documentation]:
    ///     https://support.office.com/en-ie/article/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
    ///
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.set_name(name))?;
        Ok(self.clone())
    }

    /// Freeze panes in a worksheet.
    ///
    /// The `set_freeze_panes()` method can be used to divide a worksheet into
    /// horizontal or vertical regions known as panes and to “freeze” these
    /// panes so that the splitter bars are not visible.
    ///
    /// As with Excel the split is to the top and left of the cell. So to freeze
    /// the top row and leftmost column you would use `(1, 1)` (zero-indexed).
    /// Also, you can set one of the row and col parameters as 0 if you do not
    /// want either the vertical or horizontal split. See the example below.
    ///
    /// In Excel it is also possible to set "split" panes without freezing them.
    /// That feature isn't currently supported by `rust_xlsxwriter`.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    #[wasm_bindgen(js_name = "setFreezePanes", skip_jsdoc)]
    pub fn set_freeze_panes(&self, row: xlsx::RowNum, col: xlsx::ColNum) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.set_freeze_panes(row, col))?;
        Ok(self.clone())
    }

    /// Set the top most cell in the scrolling area of a freeze pane.
    ///
    /// This method is used in conjunction with the
    /// [`Worksheet::set_freeze_panes()`] method to set the top most visible
    /// cell in the scrolling range. For example you may want to freeze the top
    /// row but have the worksheet pre-scrolled so that cell `A20` is visible in
    /// the scrolled area. See the example below.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    #[wasm_bindgen(js_name = "setFreezePanesTopCell", skip_jsdoc)]
    pub fn set_freeze_panes_top_cell(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.set_freeze_panes_top_cell(row, col))?;
        Ok(self.clone())
    }

    /// Make a worksheet the active/initially visible worksheet in a workbook.
    ///
    /// The `set_active()` method is used to specify which worksheet is
    /// initially visible in a multi-sheet workbook. If no worksheet is set then
    /// the first worksheet is made the active worksheet, like in Excel.
    ///
    /// @param {boolean} enable - Turn the property on/off. It is off by default.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setActive", skip_jsdoc)]
    pub fn set_active(&self, enable: bool) -> Worksheet {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        sheet.set_active(enable);
        self.clone()
    }

    /// Set the width for a worksheet column.
    ///
    /// The `setColumnWidth()` method is used to change the default width of a
    /// worksheet column.
    ///
    /// The ``width`` parameter sets the column width in the same units used by
    /// Excel which is: the number of characters in the default font. The
    /// default width is 8.43 in the default font of Calibri 11. The actual
    /// relationship between a string width and a column width in Excel is
    /// complex. See the [following explanation of column
    /// widths](https://support.microsoft.com/en-us/kb/214123) from the
    /// Microsoft support documentation for more details. To set the width in
    /// pixels use the {@link Worksheet#setColumnWidthPixels} method.
    ///
    /// See also the {@link Worksheet#autofit} method.
    ///
    /// @param {number} col - The zero indexed column number.
    /// @param {number} width - The column width in character units.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setColumnWidth", skip_jsdoc)]
    pub fn set_column_width(&self, col: xlsx::ColNum, width: f64) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.set_column_width(col, width))?;
        Ok(self.clone())
    }

    /// Set the width for a worksheet column in pixels.
    ///
    /// The `setColumnWidthPixels()` method is used to change the default
    /// width in pixels of a worksheet column.
    ///
    /// To set the width in Excel character units use the
    /// {Worksheet#setColumnWidth} method.
    ///
    /// See also the {@link Worksheet#autofit} method.
    ///
    /// @param {number} col - The zero indexed column number.
    /// @param {number} width - The column width in pixels.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setColumnWidthPixels", skip_jsdoc)]
    pub fn set_column_width_pixels(&self, col: xlsx::ColNum, width: u16) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.set_column_width_pixels(col, width))?;
        Ok(self.clone())
    }

    /// Set the width for a range of columns.
    ///
    /// This is a syntactic shortcut for setting the width for a range of
    /// contiguous cells. See {@link Worksheet#setColumnWidth} for more
    /// details on the single column version.
    ///
    /// @param {number} first_col - The first row of the range. Zero indexed.
    /// @param {number} last_col - The last row of the range.
    /// @param {number} width - The column width in character units.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Column exceeds Excel's worksheet
    ///   limits.
    /// - [`XlsxError::RowColumnOrderError`] - First column larger than the last
    ///   column.
    ///
    #[wasm_bindgen(js_name = "setColumnRangeWidth", skip_jsdoc)]
    pub fn set_column_range_width(
        &self,
        first_col: xlsx::ColNum,
        last_col: xlsx::ColNum,
        width: f64,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.set_column_range_width(first_col, last_col, width))?;
        Ok(self.clone())
    }

    /// Write generic data to a cell.
    ///
    /// The `write()` method writes data of type {@link ExcelData} to a worksheet.
    ///
    /// The types currently supported are:
    /// - {number}
    /// - {string}
    /// - {boolean}
    /// - {null}
    /// - {Date}
    /// - {@link Formula}
    /// - {@link Url}
    ///
    /// TODO: support bigint
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {ExcelData} data - Data to write.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    #[wasm_bindgen(js_name = "write", skip_jsdoc)]
    pub fn write(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        data: &JsExcelData,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let data: ExcelData = data.try_into()?;
        let _ = map_xlsx_error(sheet.write(row, col, data))?;
        Ok(self.clone())
    }

    /// Write formatted generic data to a cell.
    ///
    /// The `writeWithFormat()` method writes data of type {@link ExcelData} to a worksheet.
    ///
    /// See {@link Worksheet#write} for a list of supported data types.
    /// See {@link Format} for a list of supported formatting options.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {ExcelData} data - Data to write.
    /// @param {Format} format - The {@link Format} property for the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    #[wasm_bindgen(js_name = "writeWithFormat", skip_jsdoc)]
    pub fn write_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        data: &JsExcelData,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let data: ExcelData = data.try_into()?;
        let _ = map_xlsx_error(sheet.write_with_format(row, col, data, &format.lock()))?;
        Ok(self.clone())
    }

    /// Write a blank formatted worksheet cell.
    ///
    /// Write a blank cell with formatting to a worksheet cell. The format is
    /// set via a {@link Format} struct.
    ///
    /// Excel differentiates between an “Empty” cell and a “Blank” cell. An
    /// “Empty” cell is a cell which doesn’t contain data or formatting whilst a
    /// “Blank” cell doesn’t contain data but does contain formatting. Excel
    /// stores “Blank” cells but ignores “Empty” cells.
    ///
    /// The most common case for a formatted blank cell is to write a background
    /// or a border, see the example below.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {Format} format - The {@link Format} property for the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "writeBlank", skip_jsdoc)]
    pub fn write_blank(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_blank(row, col, &format.lock()))?;
        Ok(self.clone())
    }

    /// Write an unformatted string to a worksheet cell.
    ///
    /// Write an unformatted string to a worksheet cell. To write a formatted
    /// string see the {@link Worksheet#writeStringWithFormat} method below.
    ///
    /// Excel only supports UTF-8 text in the xlsx file format. Any Rust UTF-8
    /// encoded string can be written with this method. The maximum string size
    /// supported by Excel is 32,767 characters.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {string} string - The string to write to the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "writeString", skip_jsdoc)]
    pub fn write_string(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        string: &str,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_string(row, col, string))?;
        Ok(self.clone())
    }

    /// Write a formatted string to a worksheet cell.
    ///
    /// Write a string with formatting to a worksheet cell. The format is set
    /// via a {@link Format} struct which can control the font or color or
    /// properties such as bold and italic.
    ///
    /// Excel only supports UTF-8 text in the xlsx file format. Any Rust UTF-8
    /// encoded string can be written with this method. The maximum string
    /// size supported by Excel is 32,767 characters.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {string} string - The string to write to the cell.
    /// @param {Format} format - The {@link Format} property for the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "writeStringWithFormat", skip_jsdoc)]
    pub fn write_string_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        string: &str,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_string_with_format(row, col, string, &format.lock()))?;
        Ok(self.clone())
    }

    /// Write an unformatted number to a cell.
    ///
    /// Write an unformatted number to a worksheet cell. To write a formatted
    /// number see the {@link Worksheet#writeNumberWithFormat} method below.
    ///
    /// TODO: improve docs
    /// All numerical values in Excel are stored as [IEEE 754] Doubles which are
    /// the equivalent of rust's [`f64`] type. This method will accept any rust
    /// type that will convert [`Into`] a f64. These include i8, u8, i16, u16,
    /// i32, u32 and f32 but not i64 or u64, see below.
    ///
    /// IEEE 754 Doubles and f64 have around 15 digits of precision. Anything
    /// beyond that cannot be stored as a number by Excel without a loss of
    /// precision and may need to be stored as a string instead.
    ///
    /// [IEEE 754]: https://en.wikipedia.org/wiki/IEEE_754
    ///
    /// For i64/u64 you can cast the numbers `as f64` which will allow you to
    /// store the number with a loss of precision outside Excel's integer range
    /// of +/- 999,999,999,999,999 (15 digits).
    ///
    /// Excel doesn't have handling for NaN or INF floating point numbers.
    /// These will be stored as the strings "Nan", "INF", and "-INF" strings
    /// instead.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {number} number - The number to write to the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "writeNumber", skip_jsdoc)]
    pub fn write_number(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        number: f64,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_number(row, col, number))?;
        Ok(self.clone())
    }

    /// Write a formatted number to a worksheet cell.
    ///
    /// Write a number with formatting to a worksheet cell. The format is set
    /// via a {@link Format} struct which can control the numerical formatting of
    /// the number, for example as a currency or a percentage value, or the
    /// visual format, such as bold and italic text.
    ///
    /// TODO: improve docs
    /// All numerical values in Excel are stored as [IEEE 754] Doubles which are
    /// the equivalent of rust's [`f64`] type. This method will accept any rust
    /// type that will convert [`Into`] a f64. These include i8, u8, i16, u16,
    /// i32, u32 and f32 but not i64 or u64, see below.
    ///
    /// IEEE 754 Doubles and f64 have around 15 digits of precision. Anything
    /// beyond that cannot be stored as a number by Excel without a loss of
    /// precision and may need to be stored as a string instead.
    ///
    /// [IEEE 754]: https://en.wikipedia.org/wiki/IEEE_754
    ///
    /// For i64/u64 you can cast the numbers `as f64` which will allow you to
    /// store the number with a loss of precision outside Excel's integer range
    /// of +/- 999,999,999,999,999 (15 digits).
    ///
    /// Excel doesn't have handling for NaN or INF floating point numbers. These
    /// will be stored as the strings "Nan", "INF", and "-INF" strings instead.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {number} number - The number to write to the cell.
    /// @param {Format} format - The {@link Format} property for the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "writeNumberWithFormat", skip_jsdoc)]
    pub fn write_number_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        number: f64,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_number_with_format(row, col, number, &format.lock()))?;
        Ok(self.clone())
    }

    /// Write an unformatted boolean value to a cell.
    ///
    /// Write an unformatted Excel boolean value to a worksheet cell.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {boolean} boolean - The boolean value to write to the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "writeBoolean", skip_jsdoc)]
    pub fn write_boolean(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        boolean: bool,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_boolean(row, col, boolean))?;
        Ok(self.clone())
    }

    /// Write a formatted boolean value to a worksheet cell.
    ///
    /// Write a boolean value with formatting to a worksheet cell. The format is set
    /// via a {@link Format} struct which can control the numerical formatting of
    /// the number, for example as a currency or a percentage value, or the
    /// visual format, such as bold and italic text.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {boolean} boolean - The boolean value to write to the cell.
    /// @param {Format} format - The {@link Format} property for the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeBooleanWithFormat", skip_jsdoc)]
    pub fn write_boolean_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        boolean: bool,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_boolean_with_format(row, col, boolean, &format.lock()))?;
        Ok(self.clone())
    }

    /// Write an unformatted date and/or time to a worksheet cell.
    ///
    /// In general an unformatted date/time isn't very useful since a date in
    /// Excel without a format is just a number. However, this method is
    /// provided for cases where an implicit format is derived from the column
    /// or row format.
    ///
    /// However, for most use cases you should use the
    /// {@link Worksheet#writeDatetimeWithFormat} method with an explicit format.
    ///
    /// The date/time types supported are:
    /// - {Date}
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {Date} datetime - A date/time to write.
    /// @return {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "writeDatetime", skip_jsdoc)]
    pub fn write_datetime(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        datetime: &JsValue,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        if let Some(dt) = utils::datetime_of_jsval(datetime.clone()) {
            let _ = map_xlsx_error(sheet.write_datetime(row, col, dt))?;
            Ok(self.clone())
        } else {
            Err(XlsxError::InvalidDate)
        }
    }

    /// Write a formatted date and/or time to a worksheet cell.
    ///
    /// The method method writes dates/times that is of type {Date}.
    ///
    /// The date/time types supported are:
    ///- {Date}
    ///
    /// Excel stores dates and times as a floating point number with a number
    /// format to defined how it is displayed. The number format is set via a
    /// {@link Format} struct which can also control visual formatting such as bold
    /// and italic text.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {Date} datetime - A date/time to write.
    /// @param {Format} format - The {@link Format} property for the cell.
    /// @return {Worksheet} - The worksheet object.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "writeDatetimeWithFormat", skip_jsdoc)]
    pub fn write_datetime_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        datetime: &JsValue,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        if let Some(dt) = utils::datetime_of_jsval(datetime.clone()) {
            let _ = map_xlsx_error(sheet.write_datetime_with_format(row, col, dt, &format.lock()))?;
            Ok(self.clone())
        } else {
            Err(XlsxError::InvalidDate)
        }
    }

    #[wasm_bindgen(js_name = "writeFormula")]
    pub fn write_formula(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        formula: &Formula,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_formula(row, col, &*formula.lock()))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeFormulaWithFormat")]
    pub fn write_formula_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        formula: &Formula,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_formula_with_format(
            row,
            col,
            &*formula.lock(),
            &format.lock(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeUrl")]
    pub fn write_url(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        link: &Url,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_url(row, col, &*link.lock()))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeUrlWithFormat")]
    pub fn write_url_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        link: &Url,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ =
            map_xlsx_error(sheet.write_url_with_format(row, col, &*link.lock(), &format.lock()))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeUrlWithText")]
    pub fn write_url_with_text(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        link: &Url,
        text: &str,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_url_with_text(row, col, &*link.lock(), text))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeUrlWithOptions")]
    pub fn write_url_with_options(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        link: &Url,
        text: &str,
        tip: &str,
        format: Option<Format>,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_url_with_options(
            row,
            col,
            &*link.lock(),
            text,
            tip,
            format.map(|f| f.lock().clone()).as_ref(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeRichString")]
    pub fn write_rich_string(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        rich_string: &RichString,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let rich_string = rich_string.lock();
        let rich_string: Vec<_> = rich_string.iter().map(|(f, s)| (f, s.as_str())).collect();
        let _ = map_xlsx_error(sheet.write_rich_string(row, col, &rich_string))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeRichStringWithFormat")]
    pub fn write_rich_string_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        rich_string: &RichString,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let rich_string = rich_string.lock();
        let rich_string: Vec<_> = rich_string.iter().map(|(f, s)| (f, s.as_str())).collect();
        let _ = map_xlsx_error(sheet.write_rich_string_with_format(
            row,
            col,
            &rich_string,
            &format.lock(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeColumn")]
    pub fn write_column(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        values: &JsExcelDataArray,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<ExcelData> = values.try_into()?;
        let _ = map_xlsx_error(sheet.write_column(row, col, values))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeColumnWithFormat")]
    pub fn write_column_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        values: &JsExcelDataArray,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<ExcelData> = values.try_into()?;
        let _ = map_xlsx_error(sheet.write_column_with_format(row, col, values, &format.lock()))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeColumnMatrix")]
    pub fn write_column_matrix(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        data: &JsExcelDataMatrix,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<Vec<ExcelData>> = data.try_into()?;
        let _ = map_xlsx_error(sheet.write_column_matrix(row, col, values))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeRow")]
    pub fn write_row(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        values: &JsExcelDataArray,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<ExcelData> = values.try_into()?;
        let _ = map_xlsx_error(sheet.write_row(row, col, values))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeRowWithFormat")]
    pub fn write_row_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        values: &JsExcelDataArray,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<ExcelData> = values.try_into()?;
        let _ = map_xlsx_error(sheet.write_row_with_format(row, col, values, &format.lock()))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeRowMatrix")]
    pub fn write_row_matrix(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        data: &JsExcelDataMatrix,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let values: Vec<Vec<ExcelData>> = data.try_into()?;
        let _ = map_xlsx_error(sheet.write_row_matrix(row, col, values))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeArrayFormula")]
    pub fn write_array_formula(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
        formula: &Formula,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_array_formula(
            first_row,
            first_col,
            last_row,
            last_col,
            &*formula.lock(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeArrayFormulaWithFormat")]
    pub fn write_array_formula_with_format(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
        formula: &Formula,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_array_formula_with_format(
            first_row,
            first_col,
            last_row,
            last_col,
            &*formula.lock(),
            &format.lock(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeDynamicArrayFormula")]
    pub fn write_dynamic_array_formula(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
        formula: &Formula,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_dynamic_array_formula(
            first_row,
            first_col,
            last_row,
            last_col,
            &*formula.lock(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeDynamicArrayFormulaWithFormat")]
    pub fn write_dynamic_array_formula_with_format(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
        formula: &Formula,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_dynamic_array_formula_with_format(
            first_row,
            first_col,
            last_row,
            last_col,
            &*formula.lock(),
            &format.lock(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeDynamicFormula")]
    pub fn write_dyncmic_formula(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
        formula: &Formula,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_dynamic_array_formula(
            first_row,
            first_col,
            last_row,
            last_col,
            &*formula.lock(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "writeDynamicFormulaWithFormat")]
    pub fn write_dynamic_formula_with_format(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
        formula: &Formula,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.write_dynamic_array_formula_with_format(
            first_row,
            first_col,
            last_row,
            last_col,
            &*formula.lock(),
            &format.lock(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "addTable")]
    pub fn add_table(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
        table: &Table,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = sheet.add_table(first_row, first_col, last_row, last_col, &table.inner);
        Ok(self.clone())
    }

    /// Embed an image to a worksheet and fit it to a cell.
    ///
    /// This method can be used to embed a image into a worksheet cell and have
    /// the image automatically scale to the width and height of the cell. The
    /// X/Y scaling of the image is preserved but the size of the image is
    /// adjusted to fit the largest possible width or height depending on the
    /// cell dimensions.
    ///
    /// This is the equivalent of Excel's menu option to insert an image using
    /// the option to "Place in Cell" which is only available in Excel 365
    /// versions from 2023 onwards. For older versions of Excel a `#VALUE!`
    /// error is displayed.
    ///
    /// The image should be encapsulated in an {@link Image} object. See
    /// {@link Worksheet#insertImage} above for details on the supported image
    /// types.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {Image} image - The {@link Image} to insert into the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "embedImage")]
    pub fn embed_image(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        image: &Image,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.embed_image(row, col, &image.lock()))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "embedImageWithFormat")]
    pub fn embed_image_with_format(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        image: &Image,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ =
            map_xlsx_error(sheet.embed_image_with_format(row, col, &image.lock(), &format.lock()))?;
        Ok(self.clone())
    }

    /// Add an image to a worksheet.
    ///
    /// Add an image to a worksheet at a cell location. The image should be
    /// encapsulated in an {@link Image} object.
    ///
    /// The supported image formats are:
    ///
    /// - PNG
    /// - JPG
    /// - GIF: The image can be an animated gif in more recent versions of
    ///   Excel.
    /// - BMP: BMP images are only supported for backward compatibility. In
    ///   general it is best to avoid BMP images since they are not compressed.
    ///   If used, BMP images must be 24 bit, true color, bitmaps.
    ///
    /// EMF and WMF file formats will be supported in an upcoming version of the
    /// library.
    ///
    /// **NOTE on SVG files**: Excel doesn't directly support SVG files in the
    /// same way as other image file formats. It allows SVG to be inserted into
    /// a worksheet but converts them to, and displays them as, PNG files. It
    /// stores the original SVG image in the file so the original format can be
    /// retrieved. This removes the file size and resolution advantage of using
    /// SVG files. As such SVG files are not supported by `rust_xlsxwriter`
    /// since a conversion to the PNG format would be required and that format
    /// is already supported.
    ///
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {Image} image - The {@link Image} to insert into the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "insertImage", skip_jsdoc)]
    pub fn insert_image(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        image: &Image,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.insert_image(row, col, &image.lock()))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "insertImageWithOffset")]
    pub fn insert_image_with_offset(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        image: &Image,
        x_offset: u32,
        y_offset: u32,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.insert_image_with_offset(
            row,
            col,
            &image.lock(),
            x_offset,
            y_offset,
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "insertImageFitToCell")]
    pub fn insert_image_fit_to_cell(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        image: &Image,
        keep_aspect_ratio: bool,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.insert_image_fit_to_cell(
            row,
            col,
            &image.lock(),
            keep_aspect_ratio,
        ))?;
        Ok(self.clone())
    }

    /// Add an chart to a worksheet.
    ///
    /// Add a chart to a worksheet at a cell location. The chart should be
    /// encapsulated in an {@link Chart} object.
    ///
    /// The chart can be inserted as an object or as a background image.
    ///
    /// @param {number} row - The zero indexed row number.
    /// @param {number} col - The zero indexed column number.
    /// @param {Chart} chart - The {@link Chart} to insert into the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "insertChart", skip_jsdoc)]
    pub fn insert_chart(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        chart: &Chart,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.insert_chart(row, col, &chart.lock()))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "insertChartWithOffset")]
    pub fn insert_chart_with_offset(
        &self,
        row: xlsx::RowNum,
        col: xlsx::ColNum,
        chart: &Chart,
        x_offset: u32,
        y_offset: u32,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.insert_chart_with_offset(
            row,
            col,
            &chart.lock(),
            x_offset,
            y_offset,
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "clearCell")]
    pub fn clear_cell(&self, row: xlsx::RowNum, col: xlsx::ColNum) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        sheet.clear_cell(row, col);
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "clearCellFormat")]
    pub fn clear_cell_format(&self, row: xlsx::RowNum, col: xlsx::ColNum) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        sheet.clear_cell_format(row, col);
        Ok(self.clone())
    }

    /// Autofit the worksheet column widths, approximately.
    ///
    /// There is no option in the xlsx file format that can be used to say
    /// "autofit columns on loading". Auto-fitting of columns is something that
    /// Excel does at runtime when it has access to all of the worksheet
    /// information as well as the Windows functions for calculating display
    /// areas based on fonts and formatting.
    ///
    /// The `rust_xlsxwriter` library doesn't have access to the Windows
    /// functions that Excel has so it simulates autofit by calculating string
    /// widths using metrics taken from Excel.
    ///
    /// As such, there are some limitations to be aware of when using this
    /// method:
    ///
    /// - It is a simulated method and may not be accurate in all cases.
    /// - It is based on the default Excel font type and size of Calibri 11. It
    ///   will not give accurate results for other fonts or font sizes.
    /// - It doesn't take number or date formatting into account, although it
    ///   may try to in a later version.
    /// - It iterates over all the cells in a worksheet that have been populated
    ///   with data and performs a length calculation on each one, so it can
    ///   have a performance overhead for larger worksheets. See Note 1 below.
    ///
    /// This isn't perfect but for most cases it should be sufficient and if not
    /// you can adjust or prompt it by setting your own column widths via
    /// {@link Worksheet#setColumnWidth} or
    /// {@link Worksheet#setColumnWidthPixels}.
    ///
    /// The `autofit()` method ignores columns that have already been explicitly
    /// set if the width is greater than the calculated autofit width.
    /// Alternatively, setting the column width explicitly after calling
    /// `autofit()` will override the autofit value.
    ///
    /// **Note 1**: As a performance optimization when dealing with large data
    /// sets you can call `autofit()` after writing the first 50 or 100 rows.
    /// This will produce a reasonably accurate autofit for the first visible
    /// page of data without incurring the performance penalty of autofitting
    /// thousands of non-visible rows.
    #[wasm_bindgen(js_name = "autofit")]
    pub fn autofit(&self) -> Worksheet {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        sheet.autofit();
        self.clone()
    }

    #[wasm_bindgen(js_name = "autofilter")]
    pub fn autofilter(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        sheet
            .autofilter(first_row, first_col, last_row, last_col)
            .unwrap();
        Ok(self.clone())
    }

    /// Protect a worksheet from modification.
    ///
    /// The `protect()` method protects a worksheet from modification. It works
    /// by enabling a cell's `locked` and `hidden` properties, if they have been
    /// set. A **locked** cell cannot be edited and this property is on by
    /// default for all cells. A **hidden** cell will display the results of a
    /// formula but not the formula itself.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/protection_alert.png">
    ///
    /// These properties can be set using the
    /// {@link Format#setLocked}
    /// {@link Format#setUnlocked} and
    /// {@link Worksheet#setHidden} format methods. All cells
    /// have the `locked` property turned on by default (see the example below)
    /// so in general you don't have to explicitly turn it on.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "protect")]
    pub fn protect(&self) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        sheet.protect();
        Ok(self.clone())
    }

    /// Hide a worksheet.
    ///
    /// The `set_hidden()` method is used to hide a worksheet. This can be used
    /// to hide a worksheet in order to avoid confusing a user with intermediate
    /// data or calculations.
    ///
    /// In Excel a hidden worksheet can not be activated or selected so this
    /// method is mutually exclusive with the {@link Worksheet#setActive} and
    /// {@link Worksheet#setSelected} methods. In addition, since the first
    /// worksheet will default to being the active worksheet, you cannot hide
    /// the first worksheet without activating another sheet.
    ///
    /// @param {boolean} enable - Turn the property on/off. It is off by default.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self, enable: bool) -> Worksheet {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        sheet.set_hidden(enable);
        self.clone()
    }

    /// Merge a range of cells.
    ///
    /// The `mergeRange()` method allows cells to be merged together so that
    /// they act as a single area.
    ///
    /// The `mergeRange()` method writes a string to the merged cells. In order
    /// to write other data types, such as a number or a formula, you can
    /// overwrite the first cell with a call to one of the other
    /// `worksheet.write*()` functions. The same {@link Format} instance should be
    /// used as was used in the merged range, see the example below.
    ///
    /// @param {number} first_row - The first row of the range. (All zero indexed.)
    /// @param {number} first_col - The first row of the range.
    /// @param {number} last_row - The last row of the range.
    /// @param {number} last_col - The last row of the range.
    /// @param {string} value - The string to write to the cell. Other types can also be
    ///   handled. See the documentation above and the example below.
    /// @param {Format} format - The {@link Format} property for the cell.
    /// @returns {Worksheet} - The worksheet object.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::RowColumnLimitError`] - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - [`XlsxError::RowColumnOrderError`] - First row larger than the last
    ///   row.
    /// - [`XlsxError::MergeRangeSingleCell`] - A merge range cannot be a single
    ///   cell in Excel.
    /// - [`XlsxError::MergeRangeOverlaps`] - The merge range overlaps a
    ///   previous merge range.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "mergeRange", skip_jsdoc)]
    pub fn merge_range(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
        value: &str,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.merge_range(
            first_row,
            first_col,
            last_row,
            last_col,
            value,
            &format.lock(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "setRowHeight")]
    pub fn set_row_height(&mut self, row: xlsx::RowNum, height: f64) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.set_row_height(row, height))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "setRowHeightPixels")]
    pub fn set_row_height_pixels(
        &mut self,
        row: xlsx::RowNum,
        height: u16,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.set_row_height_pixels(row, height))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "setRangeWithFormat")]
    pub fn set_range_format(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.set_range_format(
            first_row,
            first_col,
            last_row,
            last_col,
            &format.lock(),
        ))?;
        Ok(self.clone())
    }

    #[wasm_bindgen(js_name = "setRangeFormatWithBorder")]
    pub fn set_range_format_with_border(
        &self,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
        format: &Format,
        border_format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut book = self.workbook.lock().unwrap();
        let sheet = book.worksheet_from_index(self.index).unwrap();
        let _ = map_xlsx_error(sheet.set_range_format_with_border(
            first_row,
            first_col,
            last_row,
            last_col,
            &format.lock(),
            &border_format.lock(),
        ))?;
        Ok(self.clone())
    }
}
