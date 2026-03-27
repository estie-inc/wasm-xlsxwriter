use crate::wrapper::Button;
use crate::wrapper::Chart;
use crate::wrapper::Color;
use crate::wrapper::DataValidation;
use crate::wrapper::ExcelDateTime;
use crate::wrapper::FilterCondition;
use crate::wrapper::Format;
use crate::wrapper::Formula;
use crate::wrapper::HeaderImagePosition;
use crate::wrapper::IgnoreError;
use crate::wrapper::Image;
use crate::wrapper::Note;
use crate::wrapper::ProtectionOptions;
use crate::wrapper::Shape;
use crate::wrapper::Sparkline;
use crate::wrapper::Table;
use crate::wrapper::Url;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Worksheet` struct represents an Excel worksheet. It handles operations
/// such as writing data to cells or formatting the worksheet layout.
#[derive(Clone)]
#[wasm_bindgen]
pub struct Worksheet {
    pub(crate) parent: Arc<Mutex<xlsx::Workbook>>,
    pub(crate) index: usize,
}

#[wasm_bindgen]
impl Worksheet {
    /// Set the worksheet name.
    ///
    /// Set the worksheet name. If no name is set the default Excel convention
    /// will be followed (Sheet1, Sheet2, etc.) in the order the worksheets are
    /// created.
    ///
    /// # Parameters
    ///
    /// - `name`: The worksheet name. It must follow the Excel rules, shown
    ///   below.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#SheetnameCannotBeBlank} - Worksheet name cannot be
    ///   blank.
    /// - {@link XlsxError#SheetnameLengthExceeded} - Worksheet name exceeds
    ///   Excel's limit of 31 characters.
    /// - {@link XlsxError#SheetnameContainsInvalidCharacter} - Worksheet name
    ///   cannot contain invalid characters: `[ ] : * ? / \`
    /// - {@link XlsxError#SheetnameStartsOrEndsWithApostrophe} - Worksheet name
    ///   cannot start or end with an apostrophe.
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_name(name)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Get the worksheet name.
    ///
    /// Get the worksheet name that was set automatically such as Sheet1,
    /// Sheet2, etc., or that was set by the user using
    /// {@link Worksheet#setName}.
    ///
    /// The worksheet name can be used to get a reference to a worksheet object
    /// using the
    /// {@link Workbook#worksheetFromName}
    /// method.
    #[wasm_bindgen(js_name = "name", skip_jsdoc)]
    pub fn name(&self) -> String {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index).unwrap().name()
    }
    /// Write an unformatted number to a cell.
    ///
    /// Write an unformatted number to a worksheet cell. To write a formatted
    /// number see the {@link Worksheet#writeNumberWithFormat} method below.
    ///
    /// All numerical values in Excel are stored as [IEEE 754] Doubles which are
    /// the equivalent of rust's `f64` type. This method will accept any rust
    /// type that will convert `Into` a f64. These include i8, u8, i16, u16,
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
    /// will be stored as the strings "NAN", "INF", and "-INF" strings or the
    /// values set with {@link Worksheet#setNanValue},
    /// {@link Worksheet#setInfinityValue} or
    /// {@link Worksheet#setNegInfinityValue}.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `number`: The number to write to the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeNumber", skip_jsdoc)]
    pub fn write_number(&self, row: u32, col: u16, number: f64) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_number(row, col, number)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a formatted number to a worksheet cell.
    ///
    /// Write a number with formatting to a worksheet cell. The format is set
    /// via a {@link Format} struct which can control the numerical formatting of
    /// the number, for example as a currency or a percentage value, or the
    /// visual format, such as bold and italic text.
    ///
    /// All numerical values in Excel are stored as [IEEE 754] Doubles which are
    /// the equivalent of rust's `f64` type. This method will accept any rust
    /// type that will convert `Into` a f64. These include i8, u8, i16, u16,
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
    /// will be stored as the strings "NAN", "INF", and "-INF" strings or the
    /// values set with {@link Worksheet#setNanValue},
    /// {@link Worksheet#setInfinityValue} or
    /// {@link Worksheet#setNegInfinityValue}.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `number`: The number to write to the cell.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeNumberWithFormat", skip_jsdoc)]
    pub fn write_number_with_format(
        &self,
        row: u32,
        col: u16,
        number: f64,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_number_with_format(row, col, number, &*format.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
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
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `string`: The string to write to the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#MaxStringLengthExceeded} - String exceeds Excel's limit
    ///   of 32,767 characters.
    #[wasm_bindgen(js_name = "writeString", skip_jsdoc)]
    pub fn write_string(&self, row: u32, col: u16, string: &str) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_string(row, col, string)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
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
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `string`: The string to write to the cell.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#MaxStringLengthExceeded} - String exceeds Excel's limit
    ///   of 32,767 characters.
    #[wasm_bindgen(js_name = "writeStringWithFormat", skip_jsdoc)]
    pub fn write_string_with_format(
        &self,
        row: u32,
        col: u16,
        string: &str,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_string_with_format(row, col, string, &*format.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write an unformatted formula to a worksheet cell.
    ///
    /// Write an unformatted Excel formula to a worksheet cell. See also the
    /// documentation on working with formulas at {@link Formula}.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `formula`: The formula to write to the cell as a string or {@link Formula}.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeFormula", skip_jsdoc)]
    pub fn write_formula(&self, row: u32, col: u16, formula: Formula) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_formula(row, col, formula.inner.lock().unwrap().clone())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a formatted formula to a worksheet cell.
    ///
    /// Write a formula with formatting to a worksheet cell. The format is set
    /// via a {@link Format} struct which can control the font or color or
    /// properties such as bold and italic.
    ///
    /// See also the documentation on working with formulas at {@link Formula}.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `formula`: The formula to write to the cell as a string or {@link Formula}.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeFormulaWithFormat", skip_jsdoc)]
    pub fn write_formula_with_format(
        &self,
        row: u32,
        col: u16,
        formula: Formula,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_formula_with_format(
                row,
                col,
                formula.inner.lock().unwrap().clone(),
                &*format.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write an  array formula to a worksheet cell.
    ///
    /// The `write_array_formula()` method writes an array formula to a
    /// cell range. In Excel an array formula is a formula that performs a
    /// calculation on a range of values. It can return a single value or a
    /// range/"array" of values.
    ///
    /// An array formula is displayed with a pair of curly brackets around the
    /// formula like this: `{=SUM(A1:B1*A2:B2)}`. The `write_array()`
    /// method doesn't require actually require these so you can omit them in
    /// the formula, and the equal sign, if you wish like this:
    /// `SUM(A1:B1*A2:B2)`.
    ///
    /// For array formulas that return a range of values you must specify the
    /// range that the return values will be written to with the `first_` and
    /// `last_` parameters. If the array formula returns a single value then the
    /// first_ and last_ parameters should be the same, as shown in the example
    /// below.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `formula`: The formula to write to the cell as a string or {@link Formula}.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row or column is larger
    ///   than the last row or column.
    #[wasm_bindgen(js_name = "writeArrayFormula", skip_jsdoc)]
    pub fn write_array_formula(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        formula: Formula,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_array_formula(
                first_row,
                first_col,
                last_row,
                last_col,
                formula.inner.lock().unwrap().clone(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a formatted array formula to a worksheet cell.
    ///
    /// Write an array formula with formatting to a worksheet cell. The format
    /// is set via a {@link Format} struct which can control the font or color or
    /// properties such as bold and italic.
    ///
    /// The `write_array()` method writes an array formula to a cell
    /// range. In Excel an array formula is a formula that performs a
    /// calculation on a range of values. It can return a single value or a
    /// range/"array" of values.
    ///
    /// An array formula is displayed with a pair of curly brackets around the
    /// formula like this: `{=SUM(A1:B1*A2:B2)}`. The `write_array()`
    /// method doesn't require actually require these so you can omit them in
    /// the formula, and the equal sign, if you wish like this:
    /// `SUM(A1:B1*A2:B2)`.
    ///
    /// For array formulas that return a range of values you must specify the
    /// range that the return values will be written to with the `first_` and
    /// `last_` parameters. If the array formula returns a single value then the
    /// first_ and last_ parameters should be the same, as shown in the example
    /// below.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `formula`: The formula to write to the cell as a string or {@link Formula}.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    #[wasm_bindgen(js_name = "writeArrayFormulaWithFormat", skip_jsdoc)]
    pub fn write_array_formula_with_format(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        formula: Formula,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_array_formula_with_format(
                first_row,
                first_col,
                last_row,
                last_col,
                formula.inner.lock().unwrap().clone(),
                &*format.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a dynamic array formula to a worksheet cell or range of cells.
    ///
    /// The `write_dynamic_array_formula()` function writes an Excel 365
    /// dynamic array formula to a cell range. Some examples of functions that
    /// return dynamic arrays are:
    ///
    /// - `FILTER()`
    /// - `RANDARRAY()`
    /// - `SEQUENCE()`
    /// - `SORTBY()`
    /// - `SORT()`
    /// - `UNIQUE()`
    /// - `XLOOKUP()`
    /// - `XMATCH()`
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `formula`: The formula to write to the cell as a string or {@link Formula}.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    #[wasm_bindgen(js_name = "writeDynamicArrayFormula", skip_jsdoc)]
    pub fn write_dynamic_array_formula(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        formula: Formula,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_dynamic_array_formula(
                first_row,
                first_col,
                last_row,
                last_col,
                formula.inner.lock().unwrap().clone(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a formatted dynamic array formula to a worksheet cell or range of
    /// cells.
    ///
    /// The `write_dynamic_array_formula_with_format()` function writes an Excel
    /// 365 dynamic array formula to a cell range. Some examples of functions
    /// that return dynamic arrays are:
    ///
    /// - `FILTER()`
    /// - `RANDARRAY()`
    /// - `SEQUENCE()`
    /// - `SORTBY()`
    /// - `SORT()`
    /// - `UNIQUE()`
    /// - `XLOOKUP()`
    /// - `XMATCH()`
    ///
    /// The format is set via a {@link Format} struct which can control the font or
    /// color or properties such as bold and italic.
    ///
    /// For array formulas that return a range of values you must specify the
    /// range that the return values will be written to with the `first_` and
    /// `last_` parameters. If the array formula returns a single value then the
    /// first_ and last_ parameters should be the same, as shown in the example
    /// below or use the {@link Worksheet#writeDynamicFormulaWithFormat}
    /// method.
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `formula`: The formula to write to the cell as a string or
    ///   {@link Formula}.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row or column is larger
    ///   than the last row or column.
    #[wasm_bindgen(js_name = "writeDynamicArrayFormulaWithFormat", skip_jsdoc)]
    pub fn write_dynamic_array_formula_with_format(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        formula: Formula,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_dynamic_array_formula_with_format(
                first_row,
                first_col,
                last_row,
                last_col,
                formula.inner.lock().unwrap().clone(),
                &*format.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a dynamic formula to a worksheet cell.
    ///
    /// The `write_dynamic_formula()` method is similar to the
    /// {@link Worksheet#writeDynamicArrayFormula} method, shown above, except
    /// that it writes a dynamic array formula to a single cell, rather than a
    /// range. This is a syntactic shortcut since the array range isn't
    /// generally known for a dynamic range and specifying the initial cell is
    /// sufficient for Excel.
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `formula`: The formula to write to the cell as a string or
    ///   {@link Formula}.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeDynamicFormula", skip_jsdoc)]
    pub fn write_dynamic_formula(
        &self,
        row: u32,
        col: u16,
        formula: Formula,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_dynamic_formula(row, col, formula.inner.lock().unwrap().clone())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a formatted dynamic formula to a worksheet cell.
    ///
    /// The `write_dynamic_formula_with_format()` method is similar to the
    /// {@link Worksheet#writeDynamicArrayFormulaWithFormat} method, shown
    /// above, except that it writes a dynamic array formula to a single cell,
    /// rather than a range. This is a syntactic shortcut since the array range
    /// isn't generally known for a dynamic range and specifying the initial
    /// cell is sufficient for Excel.
    ///
    /// For more details see the `rust_xlsxwriter` documentation section on
    /// [Dynamic Array support] and the [Dynamic array formulas] example.
    ///
    /// [Dynamic Array support]:
    ///     https://rustxlsxwriter.github.io/formulas/dynamic_arrays.html
    /// [Dynamic array formulas]:
    ///     https://rustxlsxwriter.github.io/examples/dynamic_arrays.html
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `formula`: The formula to write to the cell as a string or
    ///   {@link Formula}.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeDynamicFormulaWithFormat", skip_jsdoc)]
    pub fn write_dynamic_formula_with_format(
        &self,
        row: u32,
        col: u16,
        formula: Formula,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_dynamic_formula_with_format(
                row,
                col,
                formula.inner.lock().unwrap().clone(),
                &*format.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
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
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeBlank", skip_jsdoc)]
    pub fn write_blank(&self, row: u32, col: u16, format: &Format) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index).unwrap().write_blank(
            row,
            col,
            &*format.inner.lock().unwrap(),
        )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a url/hyperlink to a worksheet cell.
    ///
    /// Write a url/hyperlink to a worksheet cell with the default Excel
    /// "Hyperlink" cell style.
    ///
    /// There are 3 types of url/link supported by Excel:
    ///
    /// 1. Web based URIs like:
    ///
    ///    * `http://`, `https://`, `ftp://`, `ftps://` and `mailto:`.
    ///
    /// 2. Local file links using the `file://` URI.
    ///
    ///    * `file:///Book2.xlsx`
    ///    * `file:///..\Sales\Book2.xlsx`
    ///    * `file:///C:\Temp\Book1.xlsx`
    ///    * `file:///Book2.xlsx#Sheet1!A1`
    ///    * `file:///Book2.xlsx#'Sales Data'!A1:G5`
    ///
    ///    Most paths will be relative to the root folder, following the Windows
    ///    convention, so most paths should start with `file:///`. For links to
    ///    other Excel files the url string can include a sheet and cell
    ///    reference after the `"#"` anchor, as shown in the last 2 examples
    ///    above. When using Windows paths, like in the examples above, it is
    ///    best to use a Rust raw string to avoid issues with the backslashes:
    ///    `r"file:///C:\Temp\Book1.xlsx"`.
    ///
    /// 3. Internal links to a cell or range of cells in the workbook using the
    ///    pseudo-uri `internal:`:
    ///
    ///    * `internal:Sheet2!A1`
    ///    * `internal:Sheet2!A1:G5`
    ///    * `internal:'Sales Data'!A1`
    ///
    ///    Worksheet references are typically of the form `Sheet1!A1` where a
    ///    worksheet and target cell should be specified. You can also link to a
    ///    worksheet range using the standard Excel range notation like
    ///    `Sheet1!A1:B2`. Excel requires that worksheet names containing spaces
    ///    or non alphanumeric characters are single quoted as follows `'Sales
    ///    Data'!A1`.
    ///
    /// The function will escape the following characters in URLs as required by
    /// Excel, ``\s " < > \ [ ] ` ^ { }``, unless the URL already contains `%xx`
    /// style escapes. In which case it is assumed that the URL was escaped
    /// correctly by the user and will by passed directly to Excel.
    ///
    /// Excel has a limit of around 2080 characters in the url string. Strings
    /// beyond this limit will raise an error, see below.
    ///
    /// For other variants of this function see:
    ///
    /// - {@link Worksheet#writeUrlWithText} to add alternative text to the
    ///   link.
    /// - {@link Worksheet#writeUrlWithFormat} to add an alternative format to
    ///   the link.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `string`: The url string to write to the cell.
    /// - `link`: The url/hyperlink to write to the cell as a string or {@link Url}.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#MaxUrlLengthExceeded} - URL string or anchor exceeds
    ///   Excel's limit of 2080 characters.
    /// - {@link XlsxError#UnknownUrlType} - The URL has an unknown URI type. See
    ///   the supported types listed above.
    /// - {@link XlsxError#ParameterError} - {@link Url} mouseover tool tip exceeds
    ///   Excel's limit of 255 characters.
    #[wasm_bindgen(js_name = "writeUrl", skip_jsdoc)]
    pub fn write_url(&self, row: u32, col: u16, link: Url) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index).unwrap().write_url(
            row,
            col,
            link.inner.lock().unwrap().clone(),
        )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a url/hyperlink to a worksheet cell with an alternative text.
    ///
    /// Write a url/hyperlink to a worksheet cell with an alternative, user
    /// friendly, text and the default Excel "Hyperlink" cell style.
    ///
    /// This method is similar to {@link Worksheet#writeUrl}  except that you
    /// can specify an alternative string for the url. For example you could
    /// have a cell contain the link [Learn Rust](https://www.rust-lang.org)
    /// instead of the raw link <https://www.rust-lang.org>.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `link`: The url/hyperlink to write to the cell as a string or {@link Url}.
    /// - `text`: The alternative string to write to the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#MaxStringLengthExceeded} - Text string exceeds Excel's
    ///   limit of 32,767 characters.
    /// - {@link XlsxError#MaxUrlLengthExceeded} - URL string or anchor exceeds
    ///   Excel's limit of 2080 characters.
    /// - {@link XlsxError#UnknownUrlType} - The URL has an unknown URI type. See
    ///   the supported types listed above.
    /// - {@link XlsxError#ParameterError} - {@link Url} mouseover tool tip exceeds
    ///   Excel's limit of 255 characters.
    #[wasm_bindgen(js_name = "writeUrlWithText", skip_jsdoc)]
    pub fn write_url_with_text(
        &self,
        row: u32,
        col: u16,
        link: Url,
        text: &str,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_url_with_text(row, col, link.inner.lock().unwrap().clone(), text)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a url/hyperlink to a worksheet cell with a user defined format
    ///
    /// Write a url/hyperlink to a worksheet cell with a user defined format
    /// instead of the default Excel "Hyperlink" cell style.
    ///
    /// This method is similar to {@link Worksheet#writeUrl} except that you can
    /// specify an alternative format for the url.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `link`: The url/hyperlink to write to the cell as a string or {@link Url}.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#MaxUrlLengthExceeded} - URL string or anchor exceeds
    ///   Excel's limit of 2080 characters.
    /// - {@link XlsxError#UnknownUrlType} - The URL has an unknown URI type. See
    ///   the supported types listed above.
    /// - {@link XlsxError#ParameterError} - {@link Url} mouseover tool tip exceeds
    ///   Excel's limit of 255 characters.
    #[wasm_bindgen(js_name = "writeUrlWithFormat", skip_jsdoc)]
    pub fn write_url_with_format(
        &self,
        row: u32,
        col: u16,
        link: Url,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_url_with_format(
                row,
                col,
                link.inner.lock().unwrap().clone(),
                &*format.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a formatted date to a worksheet cell.
    ///
    /// The method method writes dates/times that implements {@link IntoExcelDateTime}
    /// to a worksheet cell.
    ///
    /// The date/time types supported are:
    /// - {@link ExcelDateTime}.
    ///
    /// If the `chrono` feature is enabled you can use the following types:
    ///
    /// - {@link chrono#NaiveDateTime}(https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html).
    /// - {@link chrono#NaiveDate}(https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html).
    /// - {@link chrono#NaiveTime}(https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html).
    ///
    /// If the `jiff` feature is enabled you can use the following types:
    ///
    /// - {@link jiff#civil::CivilDateTime}(https://docs.rs/jiff/latest/jiff/civil/struct.DateTime.html).
    /// - {@link jiff#civil::CivilDate}(https://docs.rs/jiff/latest/jiff/civil/struct.Date.html).
    /// - {@link jiff#civil::CivilTime}(https://docs.rs/jiff/latest/jiff/civil/struct.Time.html).
    ///
    /// Excel stores dates and times as a floating point number with a number
    /// format to defined how it is displayed. The number format is set via a
    /// {@link Format} struct which can also control visual formatting such as bold
    /// and italic text.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `date`: A date/time instance that implements {@link IntoExcelDateTime}.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeDateWithFormat", skip_jsdoc)]
    pub fn write_date_with_format(
        &self,
        row: u32,
        col: u16,
        date: &ExcelDateTime,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_date_with_format(
                row,
                col,
                &*date.inner.lock().unwrap(),
                &*format.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a formatted time to a worksheet cell.
    ///
    /// The method method writes dates/times that implements {@link IntoExcelDateTime}
    /// to a worksheet cell.
    ///
    /// The date/time types supported are:
    /// - {@link ExcelDateTime}.
    ///
    /// If the `chrono` feature is enabled you can use the following types:
    ///
    /// - {@link chrono#NaiveDateTime}(https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html).
    /// - {@link chrono#NaiveDate}(https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html).
    /// - {@link chrono#NaiveTime}(https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html).
    ///
    /// If the `jiff` feature is enabled you can use the following types:
    ///
    /// - {@link jiff#civil::CivilDateTime}(https://docs.rs/jiff/latest/jiff/civil/struct.DateTime.html).
    /// - {@link jiff#civil::CivilDate}(https://docs.rs/jiff/latest/jiff/civil/struct.Date.html).
    /// - {@link jiff#civil::CivilTime}(https://docs.rs/jiff/latest/jiff/civil/struct.Time.html).
    ///
    /// Excel stores dates and times as a floating point number with a number
    /// format to defined how it is displayed. The number format is set via a
    /// {@link Format} struct which can also control visual formatting such as bold
    /// and italic text.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `time`: A date/time instance that implements {@link IntoExcelDateTime}.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeTimeWithFormat", skip_jsdoc)]
    pub fn write_time_with_format(
        &self,
        row: u32,
        col: u16,
        time: &ExcelDateTime,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_time_with_format(
                row,
                col,
                &*time.inner.lock().unwrap(),
                &*format.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write an unformatted boolean value to a cell.
    ///
    /// Write an unformatted Excel boolean value to a worksheet cell.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `boolean`: The boolean value to write to the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeBoolean", skip_jsdoc)]
    pub fn write_boolean(&self, row: u32, col: u16, boolean: bool) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_boolean(row, col, boolean)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a formatted boolean value to a worksheet cell.
    ///
    /// Write a boolean value with formatting to a worksheet cell. The format is set
    /// via a {@link Format} struct which can control the numerical formatting of
    /// the number, for example as a currency or a percentage value, or the
    /// visual format, such as bold and italic text.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `boolean`: The boolean value to write to the cell.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "writeBooleanWithFormat", skip_jsdoc)]
    pub fn write_boolean_with_format(
        &self,
        row: u32,
        col: u16,
        boolean: bool,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .write_boolean_with_format(row, col, boolean, &*format.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Merge a range of cells.
    ///
    /// The `merge_range()` method allows cells to be merged together so that
    /// they act as a single area.
    ///
    /// The `merge_range()` method writes a string to the merged cells. In order
    /// to write other data types, such as a number or a formula, you can
    /// overwrite the first cell with a call to one of the other
    /// `worksheet.write_*()` functions. The same {@link Format} instance should be
    /// used as was used in the merged range, see the example below.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `string`: The string to write to the cell. Other types can also be
    ///   handled. See the documentation above and the example below.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    /// - {@link XlsxError#MergeRangeSingleCell} - A merge range cannot be a single
    ///   cell in Excel.
    /// - {@link XlsxError#MergeRangeOverlaps} - The merge range overlaps a
    ///   previous merge range.
    #[wasm_bindgen(js_name = "mergeRange", skip_jsdoc)]
    pub fn merge_range(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        string: &str,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index).unwrap().merge_range(
            first_row,
            first_col,
            last_row,
            last_col,
            string,
            &*format.inner.lock().unwrap(),
        )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
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
    /// **NOTE on SVG and WebP files**: Excel doesn't directly support SVG and
    /// WebP files in the same way as other image file formats. Excel allows the
    /// user to add SVG and WebP images but it converts them to PNG files and
    /// displays them in that format. As such, SVG and WebP images are not
    /// supported by `rust_xlsxwriter` since a conversion to the PNG format
    /// would be required, and that format is already supported.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The {@link Image} to insert into the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "insertImage", skip_jsdoc)]
    pub fn insert_image(&self, row: u32, col: u16, image: &Image) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_image(row, col, &*image.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add an image to a worksheet at an offset.
    ///
    /// Add an image to a worksheet at a pixel offset within a cell location.
    /// The image should be encapsulated in an {@link Image} object.
    ///
    /// This method is similar to {@link Worksheet#insertImage} except that the
    /// image can be offset from the top left of the cell.
    ///
    /// Note, it is possible to offset the image outside the target cell if
    /// required.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The {@link Image} to insert into the cell.
    /// - `x_offset`: The horizontal offset within the cell in pixels.
    /// - `y_offset`: The vertical offset within the cell in pixels.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "insertImageWithOffset", skip_jsdoc)]
    pub fn insert_image_with_offset(
        &self,
        row: u32,
        col: u16,
        image: &Image,
        x_offset: u32,
        y_offset: u32,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_image_with_offset(
                row,
                col,
                &*image.inner.lock().unwrap(),
                x_offset,
                y_offset,
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
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
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The {@link Image} to insert into the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#ParameterError} - Embedded images can only be added to
    ///   the current row in "constant memory" mode. They cannot be added to a
    ///   previously written row.
    #[wasm_bindgen(js_name = "embedImage", skip_jsdoc)]
    pub fn embed_image(&self, row: u32, col: u16, image: &Image) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index).unwrap().embed_image(
            row,
            col,
            &*image.inner.lock().unwrap(),
        )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Embed an image to a worksheet and fit it to a formatted cell.
    ///
    /// This method can be used to embed a image into a worksheet cell and have
    /// the image automatically scale to the width and height of the cell. This
    /// is similar to the {@link Worksheet#embedImage} above but it allows you
    /// to add an additional cell format using {@link Format}. This is occasionally
    /// useful if you want to center the image, set a cell border around it, or
    /// add a cell background color. See the {@link Worksheet#embedImage}
    /// example above.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The {@link Image} to insert into the cell.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "embedImageWithFormat", skip_jsdoc)]
    pub fn embed_image_with_format(
        &self,
        row: u32,
        col: u16,
        image: &Image,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .embed_image_with_format(
                row,
                col,
                &*image.inner.lock().unwrap(),
                &*format.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add an image to a worksheet and scale it fit in a cell.
    ///
    /// Add an image to a worksheet and scale it so that it fits in a cell. This
    /// is similar in effect to {@link Worksheet#embedImage} but in Excel's
    /// terminology it inserts the image placed *over* the cell instead of *in*
    /// the cell. The only advantage of this method is that the output file will
    /// work will all versions of Excel. The `Worksheet::embed_image()` method
    /// only works with versions of Excel from 2023 onwards.
    ///
    /// This method can be useful when creating a product spreadsheet with a
    /// column of images for each product. The image should be encapsulated in
    /// an {@link Image} object. See {@link Worksheet#insertImage} above for details
    /// on the supported image types. The scaling calculation for this method
    /// takes into account the DPI of the image in the same way that Excel does.
    ///
    /// There are two options, which are controlled by the `keep_aspect_ratio`
    /// parameter. The image can be scaled vertically and horizontally to occupy
    /// the entire cell or the aspect ratio of the image can be maintained so
    /// that the image is scaled to the lesser of the horizontal or vertical
    /// sizes. See the example below.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The {@link Image} to insert into the cell.
    /// - `keep_aspect_ratio`: Boolean value to maintain the aspect ratio of the
    ///   image if `true` or scale independently in the horizontal and vertical
    ///   directions if `false`.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "insertImageFitToCell", skip_jsdoc)]
    pub fn insert_image_fit_to_cell(
        &self,
        row: u32,
        col: u16,
        image: &Image,
        keep_aspect_ratio: bool,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_image_fit_to_cell(row, col, &*image.inner.lock().unwrap(), keep_aspect_ratio)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add an image to a worksheet and scale it to fit, centered in a cell.
    ///
    /// Add an image to a worksheet and scale it so that it fits, centered, in a
    /// cell. This is similar in effect to {@link Worksheet#embedImage} with a
    /// "center" cell format but in Excel's terminology it inserts the image
    /// placed *over* the cell instead of *in* the cell. The only advantage of
    /// this method is that the output file will work will all versions of
    /// Excel. The `Worksheet::embed_image()` method only works with versions of
    /// Excel from 2023 onwards.
    ///
    /// This method is similar to {@link Worksheet#insertImageFitToCell}
    /// above, except that it always keeps the aspect ratio of the image and
    /// centers the image within the cell.
    ///
    /// See the example in {@link Worksheet#embedImage} above.
    ///
    /// **Note for macOS Excel users**: the image scale and centering may appear
    /// different in Excel for macOS compared to Windows. This is an Excel
    /// issue and not a `rust_xlsxwriter` issue. See this [Microsoft support
    /// article].
    ///
    /// [Microsoft support article]: https://learn.microsoft.com/en-us/answers/questions/4938947/size-of-images-changes-(mac-windows)?forum=msoffice-all&referrer=answers
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `image`: The {@link Image} to insert into the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "insertImageFitToCellCentered", skip_jsdoc)]
    pub fn insert_image_fit_to_cell_centered(
        &self,
        row: u32,
        col: u16,
        image: &Image,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_image_fit_to_cell_centered(row, col, &*image.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Insert a background image into a worksheet.
    ///
    /// A background image can be added to a worksheet to add a watermark or
    /// display a company logo. Excel repeats the image for the entirety of the
    /// worksheet.
    ///
    /// The image should be encapsulated in an {@link Image} object. See
    /// {@link Worksheet#insertImage} above for details on the supported image
    /// types.
    ///
    /// As an alternative to background images, it should be noted that the
    /// Microsoft Excel documentation recommends setting a watermark via an
    /// image in the worksheet header. An example of that technique is shown in
    /// the {@link Worksheet#setHeaderImage} examples.
    ///
    /// # Parameters
    ///
    /// - `image`: The {@link Image} to use as the worksheet background.
    #[wasm_bindgen(js_name = "insertBackgroundImage", skip_jsdoc)]
    pub fn insert_background_image(&self, image: &Image) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_background_image(&*image.inner.lock().unwrap());
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Add a chart to a worksheet.
    ///
    /// Add a {@link Chart} to a worksheet at a cell location.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `chart`: The {@link Chart} to insert into the cell.
    ///
    /// When used with a [Chartsheet](Worksheet::new_chartsheet) the row/column
    /// arguments are ignored but it is best to use `(0, 0)` for clarity.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#ChartError} - A general error that is raised when a
    ///   chart parameter is incorrect or a chart is configured incorrectly.
    #[wasm_bindgen(js_name = "insertChart", skip_jsdoc)]
    pub fn insert_chart(&self, row: u32, col: u16, chart: &Chart) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_chart(row, col, &*chart.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add a chart to a worksheet at an offset.
    ///
    /// Add a {@link Chart} to a worksheet  at a pixel offset within a cell
    /// location.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `chart`: The {@link Chart} to insert into the cell.
    /// - `x_offset`: The horizontal offset within the cell in pixels.
    /// - `y_offset`: The vertical offset within the cell in pixels.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#ChartError} - A general error that is raised when a
    ///   chart parameter is incorrect or a chart is configured incorrectly.
    #[wasm_bindgen(js_name = "insertChartWithOffset", skip_jsdoc)]
    pub fn insert_chart_with_offset(
        &self,
        row: u32,
        col: u16,
        chart: &Chart,
        x_offset: u32,
        y_offset: u32,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_chart_with_offset(
                row,
                col,
                &*chart.inner.lock().unwrap(),
                x_offset,
                y_offset,
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add a Note to a cell.
    ///
    /// A Note is a post-it style message that is revealed when the user mouses
    /// over a worksheet cell. The presence of a Note is indicated by a small
    /// red triangle in the upper right-hand corner of the cell.
    ///
    /// In versions of Excel prior to Office 365 Notes were referred to as
    /// "Comments". The name Comment is now used for a newer style threaded
    /// comment and Note is used for the older non threaded version. See the
    /// Microsoft docs on [The difference between threaded comments and notes].
    ///
    /// [The difference between threaded comments and notes]:
    ///     https://support.microsoft.com/en-us/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3
    ///
    /// See {@link Note} for details on the properties of Notes.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `note`: The {@link Note} to insert into the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#MaxStringLengthExceeded} - Text exceeds Excel's limit of
    ///   32,713 characters.
    #[wasm_bindgen(js_name = "insertNote", skip_jsdoc)]
    pub fn insert_note(&self, row: u32, col: u16, note: &Note) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index).unwrap().insert_note(
            row,
            col,
            &*note.inner.lock().unwrap(),
        )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Insert a textbox shape into a worksheet.
    ///
    /// This method can be used to insert an Excel Textbox shape with text into
    /// a worksheet.
    ///
    /// See the {@link Shape} documentation for a detailed description of the
    /// methods that can be used to configure the size and appearance of the
    /// textbox.
    ///
    /// Note, no Excel shape other than Textbox is supported. See Support for
    /// other Excel shape
    /// types.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `shape`: The {@link Shape} to insert into the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "insertShape", skip_jsdoc)]
    pub fn insert_shape(&self, row: u32, col: u16, shape: &Shape) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_shape(row, col, &*shape.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Insert a textbox shape into a worksheet cell at an offset.
    ///
    /// This method can be used to insert an Excel Textbox shape with text into
    /// a worksheet cell at a pixel offset.
    ///
    /// See the {@link Shape} documentation for a detailed description of the
    /// methods that can be used to configure the size and appearance of the
    /// textbox.
    ///
    /// Note, no Excel shape other than Textbox is supported. See Support for
    /// other Excel shape
    /// types.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `shape`: The {@link Shape} to insert into the cell.
    /// - `x_offset`: The horizontal offset within the cell in pixels.
    /// - `y_offset`: The vertical offset within the cell in pixels.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "insertShapeWithOffset", skip_jsdoc)]
    pub fn insert_shape_with_offset(
        &self,
        row: u32,
        col: u16,
        shape: &Shape,
        x_offset: u32,
        y_offset: u32,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_shape_with_offset(
                row,
                col,
                &*shape.inner.lock().unwrap(),
                x_offset,
                y_offset,
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Make all worksheet notes visible when the file loads.
    ///
    /// By default Excel hides cell notes until the user mouses over the parent
    /// cell. However, if required you can make all worksheet notes visible when
    /// the worksheet loads. You can also make individual notes visible using
    /// the {@link Note#setVisible} method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "showAllNotes", skip_jsdoc)]
    pub fn show_all_notes(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .show_all_notes(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the default author name for all the notes in the worksheet.
    ///
    /// The Note author is the creator of the note. In Excel the author name is
    /// taken from the user name in the options/preference dialog. The note
    /// author name appears in two places: at the start of the note text in bold
    /// and at the bottom of the worksheet in the status bar.
    ///
    /// If no name is specified the default name "Author" will be applied to the
    /// note. The author name for individual notes can be set via the
    /// {@link Note#setAuthor} method. Alternatively
    /// this method can be used to set the default author name for all notes in
    /// a worksheet.
    ///
    /// # Parameters
    ///
    /// - `name`: The note author name. Must be less than or equal to the Excel
    ///   limit of 52 characters.
    #[wasm_bindgen(js_name = "setDefaultNoteAuthor", skip_jsdoc)]
    pub fn set_default_note_author(&self, name: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_default_note_author(name);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Add a Excel Form Control button object to a worksheet.
    ///
    /// Add a {@link Button} to a worksheet at a cell location. The worksheet button
    /// object is mainly provided as a way of triggering a VBA macro, see
    /// Working with VBA macros for more details.
    ///
    /// Note, Button is the only VBA Control supported by `rust_xlsxwriter`. It
    /// is unlikely that any other Excel form elements will be added in the
    /// future due to the implementation effort required.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `button`: The {@link Button} to insert into the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "insertButton", skip_jsdoc)]
    pub fn insert_button(&self, row: u32, col: u16, button: &Button) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_button(row, col, &*button.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add a Excel Form Control button object to a  at an offset.
    ///
    /// Add a {@link Button} to a worksheet  at a pixel offset within a cell
    /// location. See {@link Worksheet#insertButton} above
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `button`: The {@link Button} to insert into the cell.
    /// - `x_offset`: The horizontal offset within the cell in pixels.
    /// - `y_offset`: The vertical offset within the cell in pixels.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "insertButtonWithOffset", skip_jsdoc)]
    pub fn insert_button_with_offset(
        &self,
        row: u32,
        col: u16,
        button: &Button,
        x_offset: u32,
        y_offset: u32,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_button_with_offset(
                row,
                col,
                &*button.inner.lock().unwrap(),
                x_offset,
                y_offset,
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Insert a boolean checkbox in a worksheet cell.
    ///
    /// Checkboxes are a [new feature] added to Excel in 2024. They are a way of
    /// displaying a boolean value as a checkbox in a cell. The underlying value
    /// is still an Excel `TRUE/FALSE` boolean value and can be used in formulas
    /// and in references.
    ///
    /// [new feature]:
    ///     https://techcommunity.microsoft.com/blog/excelblog/introducing-checkboxes-in-excel/4173561
    ///
    /// The `insert_checkbox()` method can be used to replicate this behavior,
    /// see the examples below.
    ///
    /// The checkbox feature is only available in Excel versions from 2024 and
    /// later. In older versions the value will be displayed as a standard Excel
    /// `TRUE` or `FALSE` boolean. In fact Excel actually stores a checkbox as a
    /// normal boolean but with a special format. If required you can make use
    /// of this property to create a checkbox with
    /// {@link Worksheet#writeBooleanWithFormat} and a cell format that has
    /// the {@link Format#setCheckbox} property set, see the second example
    /// below.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `boolean`: The boolean value to display as a checkbox.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "insertCheckbox", skip_jsdoc)]
    pub fn insert_checkbox(&self, row: u32, col: u16, boolean: bool) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_checkbox(row, col, boolean)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Insert a boolean checkbox in a worksheet cell with a cell format.
    ///
    /// This method allow you to insert a boolean checkbox in a worksheet cell
    /// with a background color or other cell format property.
    ///
    /// See the {@link Worksheet#insertCheckbox} method above for more details.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `boolean`: The boolean value to display as a checkbox.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "insertCheckboxWithFormat", skip_jsdoc)]
    pub fn insert_checkbox_with_format(
        &self,
        row: u32,
        col: u16,
        boolean: bool,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .insert_checkbox_with_format(row, col, boolean, &*format.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the height for a row of cells.
    ///
    /// The `set_row_height()` method is used to change the default height of a
    /// row. The height is specified in character units, where the default
    /// height is 15. Excel allows height values in increments of 0.25.
    ///
    /// It is generally preferable to set the height in pixels using the
    /// {@link Worksheet#setRowHeightPixels} method since it is more
    /// precise.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `height`: The row height in character units.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row exceeds Excel's worksheet
    ///   limits.
    #[wasm_bindgen(js_name = "setRowHeight", skip_jsdoc)]
    pub fn set_row_height(&self, row: u32, height: f64) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_row_height(row, height)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the height for a row of cells, in pixels.
    ///
    /// The `set_row_height_pixels()` method is used to change the default height of a
    /// row. The height is specified in pixels, where the default
    /// height is 20.
    ///
    /// To specify the height in Excel's character units use the
    /// {@link Worksheet#setRowHeight} method.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `height`: The row height in pixels.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row exceeds Excel's worksheet
    ///   limits.
    #[wasm_bindgen(js_name = "setRowHeightPixels", skip_jsdoc)]
    pub fn set_row_height_pixels(&self, row: u32, height: u32) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_row_height_pixels(row, height)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the format for a row of cells.
    ///
    /// The `set_row_format()` method is used to change the default format of a
    /// row. Any unformatted data written to that row will then adopt that
    /// format. Formatted data written to the row will maintain its own cell
    /// format. See the example below.
    ///
    /// A future version of this library may support automatic merging of
    /// explicit cell formatting with the row formatting but that isn't
    /// currently supported.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row exceeds Excel's worksheet
    ///   limits.
    #[wasm_bindgen(js_name = "setRowFormat", skip_jsdoc)]
    pub fn set_row_format(&self, row: u32, format: &Format) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_row_format(row, &*format.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Group a range of rows into a worksheet outline group.
    ///
    /// In Excel an outline is a group of rows or columns that can be collapsed
    /// or expanded to simplify hierarchical data. It is most often used with
    /// the `SUBTOTAL()` function. See the examples below and the the
    /// documentation on [Grouping and outlining
    /// data](../worksheet/index.html#grouping-and-outlining-data).
    ///
    /// A grouping is created as follows:
    ///
    /// Which creates a grouping at level 1:
    ///
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_rows_intro1.png">
    ///
    /// Hierarchical sub-groups are created by repeating the method calls for a
    /// sub-range of an upper level group:
    ///
    /// This creates the following grouping and sub-grouping at levels 1 and 2:
    ///
    /// src="https://rustxlsxwriter.github.io/images/worksheet_group_rows_intro2.png">
    ///
    /// It should be noted that Excel requires outline groups at the same level
    /// to be separated by at least one row (or column) or else it will merge
    /// them into a single group. This is generally to allow a subtotal
    /// row/column.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. Zero indexed.
    /// - `last_row`: The last row of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row exceeds Excel's worksheet
    ///   limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row. Note, to reverse the group direction see the
    ///   {@link Worksheet#groupSymbolsAbove} method.
    /// - {@link XlsxError#MaxGroupLevelExceeded} - Group depth level exceeds
    ///   Excel's limit of 8 levels.
    #[wasm_bindgen(js_name = "groupRows", skip_jsdoc)]
    pub fn group_rows(&self, first_row: u32, last_row: u32) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .group_rows(first_row, last_row)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Group a range of rows into a collapsed worksheet outline group.
    ///
    /// In Excel an outline is a group of rows or columns that can be collapsed
    /// or expanded to simplify hierarchical data. It is most often used with
    /// the `SUBTOTAL()` function. See the examples below and the the
    /// documentation on [Grouping and outlining
    /// data](../worksheet/index.html#grouping-and-outlining-data).
    ///
    /// See {@link Worksheet#groupRows} above for an explanation on how to create
    /// sub-groupings.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. Zero indexed.
    /// - `last_row`: The last row of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row exceeds Excel's worksheet
    ///   limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row. Note, to reverse the group direction see the
    ///   {@link Worksheet#groupSymbolsAbove} method.
    /// - {@link XlsxError#MaxGroupLevelExceeded} - Group depth level exceeds
    ///   Excel's limit of 8 levels.
    #[wasm_bindgen(js_name = "groupRowsCollapsed", skip_jsdoc)]
    pub fn group_rows_collapsed(&self, first_row: u32, last_row: u32) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .group_rows_collapsed(first_row, last_row)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Group a range of columns into a worksheet outline group.
    ///
    /// In Excel an outline is a group of rows or columns that can be collapsed
    /// or expanded to simplify hierarchical data. It is most often used with
    /// the `SUBTOTAL()` function. See the examples below and the the
    /// documentation on [Grouping and outlining
    /// data](../worksheet/index.html#grouping-and-outlining-data).
    ///
    /// See {@link Worksheet#groupRows} above for an explanation on how to create
    /// sub-groupings.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row exceeds Excel's worksheet
    ///   limits.
    /// - {@link XlsxError#RowColumnOrderError} - First column greater than the last
    ///   column. Note, to reverse the group direction see the
    ///   {@link Worksheet#groupSymbolsToLeft} method.
    /// - {@link XlsxError#MaxGroupLevelExceeded} - Group depth level exceeds
    ///   Excel's limit of 8 levels.
    #[wasm_bindgen(js_name = "groupColumns", skip_jsdoc)]
    pub fn group_columns(&self, first_col: u16, last_col: u16) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .group_columns(first_col, last_col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Group a range of columns into a collapsed worksheet outline group.
    ///
    /// In Excel an outline is a group of rows or columns that can be collapsed
    /// or expanded to simplify hierarchical data. It is most often used with
    /// the `SUBTOTAL()` function. See the examples below and the the
    /// documentation on [Grouping and outlining
    /// data](../worksheet/index.html#grouping-and-outlining-data).
    ///
    /// See {@link Worksheet#groupRows} above for an explanation on how to
    /// create sub-groupings.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row exceeds Excel's worksheet
    ///   limits.
    /// - {@link XlsxError#RowColumnOrderError} - First column greater than the
    ///   last column. Note, to reverse the group direction see the
    ///   {@link Worksheet#groupSymbolsToLeft} method.
    /// - {@link XlsxError#MaxGroupLevelExceeded} - Group depth level exceeds
    ///   Excel's limit of 7 levels.
    #[wasm_bindgen(js_name = "groupColumnsCollapsed", skip_jsdoc)]
    pub fn group_columns_collapsed(&self, first_col: u16, last_col: u16) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .group_columns_collapsed(first_col, last_col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Place the row outline group expand/collapse symbols above the range.
    ///
    /// This method toggles the Excel worksheet option to place the outline
    /// group expand/collapse symbols `[+]` and `[-]` above the group ranges
    /// instead of below for row ranges.
    ///
    /// In Excel this is a worksheet wide option and will apply to all row
    /// outlines in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "groupSymbolsAbove", skip_jsdoc)]
    pub fn group_symbols_above(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .group_symbols_above(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Place the column outline group expand/collapse symbols to the left of
    /// the range.
    ///
    /// This method toggles the Excel worksheet option to place the outline
    /// group expand/collapse symbols `[+]` and `[-]` to the left of the group
    /// ranges instead of to the right, for column ranges.
    ///
    /// In Excel this is a worksheet wide option and will apply to all column
    /// outlines in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "groupSymbolsToLeft", skip_jsdoc)]
    pub fn group_symbols_to_left(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .group_symbols_to_left(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Hide a worksheet row.
    ///
    /// The `set_row_hidden()` method is used to hide a row. This can be
    /// used, for example, to hide intermediary steps in a complicated
    /// calculation.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row exceeds Excel's worksheet
    ///   limits.
    #[wasm_bindgen(js_name = "setRowHidden", skip_jsdoc)]
    pub fn set_row_hidden(&self, row: u32) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_row_hidden(row)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Unhide a user hidden worksheet row.
    ///
    /// The `set_row_unhidden()` method is used to unhide a previously hidden
    /// row. This can occasionally be useful when used in conjunction with
    /// autofilter rules.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row exceeds Excel's worksheet
    ///   limits.
    #[wasm_bindgen(js_name = "setRowUnhidden", skip_jsdoc)]
    pub fn set_row_unhidden(&self, row: u32) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_row_unhidden(row)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the default row height for all rows in a worksheet, efficiently.
    ///
    /// This method can be used to efficiently set the default row height for
    /// all rows in a worksheet. It is efficient because it uses an Excel
    /// optimization to adjust the row heights with a single XML element. By
    /// contrast, using {@link Worksheet#setRowHeight} for every row in a
    /// worksheet would require a million XML elements and would result in a
    /// very large file.
    ///
    /// The height is specified in character units, where the default height is
    /// 15. Excel allows height values in increments of 0.25.
    ///
    /// Individual row heights can be set via {@link Worksheet#setRowHeight}.
    ///
    /// Note, there is no equivalent method for columns because the file format
    /// already optimizes the storage of a large number of contiguous columns.
    ///
    /// # Parameters
    ///
    /// - `height`: The row height in character units. Must be greater than 0.0.
    #[wasm_bindgen(js_name = "setDefaultRowHeight", skip_jsdoc)]
    pub fn set_default_row_height(&self, height: f64) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_default_row_height(height);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the default row height in pixels for all rows in a worksheet,
    /// efficiently.
    ///
    /// The `set_default_row_height_pixels()` method is used to change the
    /// default height of all rows in a worksheet. See
    /// {@link Worksheet#setDefaultRowHeight} above for an explanation.
    ///
    /// The height is specified in pixels, where the default height is 20.
    ///
    /// # Parameters
    ///
    /// - `height`: The row height in pixels.
    #[wasm_bindgen(js_name = "setDefaultRowHeightPixels", skip_jsdoc)]
    pub fn set_default_row_height_pixels(&self, height: u32) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_default_row_height_pixels(height);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Hide all unused rows in a worksheet, efficiently.
    ///
    /// This method can be used to efficiently hide unused rows in a worksheet.
    /// It is efficient because it uses an Excel optimization to hide the rows
    /// with a single XML element. By contrast, using
    /// {@link Worksheet#setRowHidden} for the majority of rows in a worksheet
    /// would require a million XML elements and would result in a very large
    /// file.
    ///
    /// "Unused" in this context means that the row doesn't contain data,
    /// formatting, or any changes such as the row height.
    ///
    /// To hide individual rows use the {@link Worksheet#setRowHidden} method.
    ///
    /// Note, there is no equivalent method for columns because the file format
    /// already optimizes the storage of a large number of contiguous columns.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "hideUnusedRows", skip_jsdoc)]
    pub fn hide_unused_rows(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .hide_unused_rows(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the width for a worksheet column.
    ///
    /// The `set_column_width()` method is used to change the default width of a
    /// worksheet column.
    ///
    /// The `width` parameter sets the column width in the same units used by
    /// Excel which is: the number of characters in the default font. The
    /// default width is 8.43 in the default font of Calibri 11. The actual
    /// relationship between a string width and a column width in Excel is
    /// complex. See the [following explanation of column
    /// widths](https://support.microsoft.com/en-us/kb/214123) from the
    /// Microsoft support documentation for more details.
    ///
    /// It is generally preferable to set the width in pixels using the
    /// {@link Worksheet#setColumnWidthPixels} method since it is more
    /// precise.
    ///
    /// See also the {@link Worksheet#autofit} method.
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    /// - `width`: The column width in character units.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Column exceeds Excel's worksheet
    ///   limits.
    #[wasm_bindgen(js_name = "setColumnWidth", skip_jsdoc)]
    pub fn set_column_width(&self, col: u16, width: f64) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_column_width(col, width)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the width for a worksheet column in pixels.
    ///
    /// The `set_column_width_pixels()` method is used to change the default
    /// width in pixels of a worksheet column.
    ///
    /// To set the width in Excel character units use the
    /// {@link Worksheet#setColumnWidth} method.
    ///
    /// See also the {@link Worksheet#autofit} method.
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    /// - `width`: The column width in pixels.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Column exceeds Excel's worksheet
    ///   limits.
    #[wasm_bindgen(js_name = "setColumnWidthPixels", skip_jsdoc)]
    pub fn set_column_width_pixels(&self, col: u16, width: u32) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_column_width_pixels(col, width)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the column pixel width in auto-fit mode.
    ///
    /// In Excel the width of an auto-fitted column will increase if the user
    /// edits a number and increases the number of digits past the previous
    /// maximum width. This behavior doesn't apply to strings or when the number
    /// of digits is decreased. It also doesn't apply to columns that have been
    /// set manually.
    ///
    /// The `Worksheet::set_column_autofit_width()` method emulates this auto-fit
    /// behavior whereas the {@link Worksheet#setColumnWidthPixels} method,
    /// see above, is equivalent to setting the width manually.
    ///
    /// The distinction is subtle and most users are unaware of this behavior in
    /// Excel. However, it is supported for users who wish to implement their
    /// own version of auto-fit.
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    /// - `width`: The column width in pixels.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Column exceeds Excel's worksheet
    ///   limits.
    #[wasm_bindgen(js_name = "setColumnAutofitWidth", skip_jsdoc)]
    pub fn set_column_autofit_width(&self, col: u16, width: u32) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_column_autofit_width(col, width)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the format for a column of cells.
    ///
    /// The `set_column_format()` method is used to change the default format of a
    /// column. Any unformatted data written to that column will then adopt that
    /// format. Formatted data written to the column will maintain its own cell
    /// format. See the example below.
    ///
    /// A future version of this library may support automatic merging of
    /// explicit cell formatting with the column formatting but that isn't
    /// currently supported.
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Column exceeds Excel's worksheet
    ///   limits.
    #[wasm_bindgen(js_name = "setColumnFormat", skip_jsdoc)]
    pub fn set_column_format(&self, col: u16, format: &Format) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_column_format(col, &*format.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Hide a worksheet column.
    ///
    /// The `set_column_hidden()` method is used to hide a column. This can be
    /// used, for example, to hide intermediary steps in a complicated
    /// calculation.
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Column exceeds Excel's worksheet
    ///   limits.
    #[wasm_bindgen(js_name = "setColumnHidden", skip_jsdoc)]
    pub fn set_column_hidden(&self, col: u16) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_column_hidden(col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the width for a range of columns.
    ///
    /// This is a syntactic shortcut for setting the width for a range of
    /// contiguous cells. See {@link Worksheet#setColumnWidth} for more
    /// details on the single column version.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    /// - `width`: The column width in character units.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Column exceeds Excel's worksheet
    ///   limits.
    /// - {@link XlsxError#RowColumnOrderError} - First column larger than the last
    ///   column.
    #[wasm_bindgen(js_name = "setColumnRangeWidth", skip_jsdoc)]
    pub fn set_column_range_width(
        &self,
        first_col: u16,
        last_col: u16,
        width: f64,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_column_range_width(first_col, last_col, width)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the pixel width for a range of columns.
    ///
    /// This is a syntactic shortcut for setting the width in pixels for a range of
    /// contiguous cells. See {@link Worksheet#setColumnWidthPixels} for more
    /// details on the single column version.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    /// - `width`: The column width in pixels.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Column exceeds Excel's worksheet
    ///   limits.
    /// - {@link XlsxError#RowColumnOrderError} - First column larger than the last
    ///   column.
    #[wasm_bindgen(js_name = "setColumnRangeWidthPixels", skip_jsdoc)]
    pub fn set_column_range_width_pixels(
        &self,
        first_col: u16,
        last_col: u16,
        width: u32,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_column_range_width_pixels(first_col, last_col, width)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the format for a range of columns.
    ///
    /// This is a syntactic shortcut for setting the format for a range of
    /// contiguous cells. See {@link Worksheet#setColumnFormat} for more
    /// details on the single column version.
    ///
    /// Note, this method can be used to set the cell format for the entire
    /// worksheet when applied to all the columns in the worksheet (see the
    /// example below). This is relatively efficient since it is stored as a
    /// single XML element. This is also how Excel applies a format to an entire
    /// worksheet.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Column exceeds Excel's worksheet
    ///   limits.
    /// - {@link XlsxError#RowColumnOrderError} - First column larger than the last
    ///   column.
    #[wasm_bindgen(js_name = "setColumnRangeFormat", skip_jsdoc)]
    pub fn set_column_range_format(
        &self,
        first_col: u16,
        last_col: u16,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_column_range_format(first_col, last_col, &*format.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Hide a range of worksheet columns.
    ///
    /// This is a syntactic shortcut for hiding a range of contiguous cells. See
    /// {@link Worksheet#setColumnHidden} for more details on the single column
    /// version.
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. Zero indexed.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Column exceeds Excel's worksheet
    ///   limits.
    /// - {@link XlsxError#RowColumnOrderError} - First column larger than the last
    ///   column.
    #[wasm_bindgen(js_name = "setColumnRangeHidden", skip_jsdoc)]
    pub fn set_column_range_hidden(&self, first_col: u16, last_col: u16) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_column_range_hidden(first_col, last_col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the default cell format for a worksheet.
    ///
    /// See the {@link Workbook#setDefaultFormat} method for details.
    ///
    /// This method is only required when a worksheet is created independently
    /// of a workbook but you still wish to change the default format. The
    /// format must also be changed via {@link Workbook#setDefaultFormat} using
    /// the same format and dimensions in the parent workbook.
    ///
    /// {@link Workbook#setDefaultFormat}: crate::workbook::Workbook::set_default_format
    ///
    /// # Parameters
    ///
    /// - `format`: The new default {@link Format} property for the worksheet.
    /// - `row_height`: The default row height in pixels.
    /// - `col_width`: The default column width in pixels. Only fonts that have
    ///   the following column pixel width are supported: 56, 64, 72, 80, 96,
    ///   104 and 120.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#DefaultFormatError} - This error occurs if you try to
    ///   set the default format after formats have been added to a worksheet or
    ///   if the the pixel column width is one of the supported values shown
    ///   above.
    #[wasm_bindgen(js_name = "setDefaultFormat", skip_jsdoc)]
    pub fn set_default_format(
        &self,
        format: &Format,
        row_height: u32,
        col_width: u32,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_default_format(&*format.inner.lock().unwrap(), row_height, col_width)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the autofilter area in the worksheet.
    ///
    /// The `autofilter()` method allows an autofilter to be added to a
    /// worksheet. An autofilter is a way of adding drop down lists to the
    /// headers of a 2D range of worksheet data. This allows users to filter the
    /// data based on simple criteria so that some data is shown and some is
    /// hidden.
    ///
    /// See the {@link Worksheet#filterColumn} method for an explanation of how to
    /// set a filter conditions for columns in the autofilter range.
    ///
    /// Note, Excel only allows one autofilter range per worksheet so calling
    /// this method multiple times will overwrite the previous range.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    /// - {@link XlsxError#AutofilterRangeOverlaps} - The autofilter range overlaps
    ///   a table autofilter range.
    #[wasm_bindgen(js_name = "autofilter", skip_jsdoc)]
    pub fn autofilter(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .autofilter(first_row, first_col, last_row, last_col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the filter condition for a column in an autofilter range.
    ///
    /// The {@link Worksheet#autofilter} method sets the cell range for an
    /// autofilter but in order to filter rows within the filter area you must
    /// also add a filter condition.
    ///
    /// Excel supports two main types of filter. The first, and most common, is
    /// a list filter where the user selects the items to filter from a list of
    /// all the values in the the column range:
    ///
    /// The other main type of filter is a custom filter where the user can
    /// specify 1 or 2 conditions like ">= 4000" and "<= 6000":
    ///
    /// src="https://rustxlsxwriter.github.io/images/autofilter_custom.png">
    ///
    /// In Excel these are mutually exclusive and you will need to choose one or
    /// the other via the {@link FilterCondition} struct parameter.
    ///
    /// For more details on setting filter conditions see {@link FilterCondition}
    /// and the [Working with Autofilters] section of the Users Guide.
    ///
    /// [Working with Autofilters]:
    ///     https://rustxlsxwriter.github.io/formulas/autofilters.html
    ///
    /// Note, there are some limitations on autofilter conditions. The main one
    /// is that the hiding of rows that don't match a filter is not an automatic
    /// part of the file format. Instead it is necessary to hide rows that don't
    /// match the filters. The `rust_xlsxwriter` library does this automatically
    /// and in most cases will get it right, however, there may be cases where
    /// you need to manually hide some of the rows. See [Auto-hiding filtered
    /// rows].
    ///
    /// [Auto-hiding filtered rows]:
    ///     https://rustxlsxwriter.github.io/formulas/autofilters.html#auto-hiding-filtered-rows
    ///
    /// # Parameters
    ///
    /// - `col`: The zero indexed column number.
    /// - `filter_condition`: The column filter condition defined by the
    ///   {@link FilterCondition} struct.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Column exceeds Excel's worksheet
    ///   limits.
    /// - {@link XlsxError#ParameterError} - Parameter error for the following
    ///   issues:
    ///   - The {@link Worksheet#autofilter} range hasn't been set.
    ///   - The column is outside the {@link Worksheet#autofilter} range.
    ///   - The {@link FilterCondition} doesn't have a condition set.
    #[wasm_bindgen(js_name = "filterColumn", skip_jsdoc)]
    pub fn filter_column(
        &self,
        col: u16,
        filter_condition: &FilterCondition,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .filter_column(col, &*filter_condition.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Turn off the option to automatically hide rows that don't match filters.
    ///
    /// Rows that don't match autofilter conditions are hidden by Excel at
    /// runtime. This feature isn't an automatic part of the file format and in
    /// practice it is necessary for the user to hide rows that don't match the
    /// applied filters. The `rust_xlsxwriter` library tries to do this
    /// automatically and in most cases will get it right, however, there may be
    /// cases where you need to manually hide some of the rows and may want to
    /// turn off the automatic handling using `filter_automatic_off()`.
    ///
    /// See [Auto-hiding filtered rows] in the User Guide.
    ///
    /// [Auto-hiding filtered rows]:
    ///     https://rustxlsxwriter.github.io/formulas/autofilters.html#auto-hiding-filtered-rows
    #[wasm_bindgen(js_name = "filterAutomaticOff", skip_jsdoc)]
    pub fn filter_automatic_off(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .filter_automatic_off();
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Add a table to a worksheet.
    ///
    /// Tables in Excel are a way of grouping a range of cells into a single
    /// entity that has common formatting or that can be referenced from
    /// formulas. Tables can have column headers, autofilters, total rows,
    /// column formulas and different formatting styles.
    ///
    /// The headers and total row of a table should be configured via a
    /// {@link Table} struct but the table data can be added via standard
    /// {@link Worksheet#write} methods.
    ///
    /// For more information on tables see the Microsoft documentation on
    /// [Overview of Excel tables].
    ///
    /// [Overview of Excel tables]:
    ///     https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    /// Note, you need to ensure that the `first_row` and `last_row` range
    /// includes all the rows for the table including the header and the total
    /// row, if present.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    /// - {@link XlsxError#TableError} - A general error that is raised when a
    ///   table parameter is incorrect or a table is configured incorrectly.
    /// - {@link XlsxError#TableRangeOverlaps} - The table range overlaps a
    ///   previous table range.
    /// - {@link XlsxError#AutofilterRangeOverlaps} - The table autofilter range
    ///   overlaps the worksheet autofilter range.
    #[wasm_bindgen(js_name = "addTable", skip_jsdoc)]
    pub fn add_table(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        table: &Table,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index).unwrap().add_table(
            first_row,
            first_col,
            last_row,
            last_col,
            &*table.inner.lock().unwrap(),
        )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add a data validation to one or more cells to restrict user input based
    /// on types and rules.
    ///
    /// Data validation is a feature of Excel which allows you to restrict the
    /// data that a user enters in a cell and to display associated help and
    /// warning messages. It also allows you to restrict input to values in a
    /// dropdown list.
    ///
    /// See {@link DataValidation} for more information and examples.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `data_validation`: A {@link DataValidation} data validation instance.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    #[wasm_bindgen(js_name = "addDataValidation", skip_jsdoc)]
    pub fn add_data_validation(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        data_validation: &DataValidation,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .add_data_validation(
                first_row,
                first_col,
                last_row,
                last_col,
                &*data_validation.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add a sparkline to a worksheet cell.
    ///
    /// Sparklines are a feature of Excel 2010+ which allows you to add small
    /// charts to worksheet cells. These are useful for showing data trends in a
    /// compact visual format.
    ///
    /// The `add_sparkline()` method allows you to add a sparkline to a single
    /// cell that displays data from a 1D range of cells.
    ///
    /// The sparkline can be configured with all the parameters supported by
    /// Excel. See {@link Sparkline} for details.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `sparkline`: The {@link Sparkline} to insert into the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#SparklineError} - An error that is raised when there is
    ///   an parameter error with the sparkline.
    /// - {@link XlsxError#ChartError} - An error that is raised when there is an
    ///   parameter error with the data range for the sparkline.
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#SheetnameCannotBeBlank} - Worksheet name in chart range
    ///   cannot be blank.
    /// - {@link XlsxError#SheetnameLengthExceeded} - Worksheet name in chart range
    ///   exceeds Excel's limit of 31 characters.
    /// - {@link XlsxError#SheetnameContainsInvalidCharacter} - Worksheet name in
    ///   chart range cannot contain invalid characters: `[ ] : * ? / \`
    /// - {@link XlsxError#SheetnameStartsOrEndsWithApostrophe} - Worksheet name in
    ///   chart range cannot start or end with an apostrophe.
    #[wasm_bindgen(js_name = "addSparkline", skip_jsdoc)]
    pub fn add_sparkline(
        &self,
        row: u32,
        col: u16,
        sparkline: &Sparkline,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .add_sparkline(row, col, &*sparkline.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add a sparkline group to a worksheet range.
    ///
    /// Sparklines are a feature of Excel 2010+ which allows you to add small
    /// charts to worksheet cells. These are useful for showing data trends in a
    /// compact visual format.
    ///
    /// In Excel sparklines can be added as a single entity in a cell that
    /// refers to a 1D data range or as a "group" sparkline that is applied
    /// across a 1D range and refers to data in a 2D range. A grouped sparkline
    /// uses one sparkline for the specified range and any changes to it are
    /// applied to the entire sparkline group.
    ///
    /// The {@link Worksheet#addSparkline} method shown above allows you to add
    /// a sparkline to a single cell that displays data from a 1D range of cells
    /// whereas `add_sparkline_group()` applies the group sparkline to a range.
    ///
    /// The sparkline can be configured with all the parameters supported by
    /// Excel. See {@link Sparkline} for details.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `sparkline`: The {@link Sparkline} to insert into the cell.
    ///
    /// # Errors
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#SparklineError} - An error that is raised when there is
    ///   an parameter error with the sparkline.
    /// - {@link XlsxError#ChartError} - An error that is raised when there is an
    ///   parameter error with the data range for the sparkline.
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#SheetnameCannotBeBlank} - Worksheet name in chart range
    ///   cannot be blank.
    /// - {@link XlsxError#SheetnameLengthExceeded} - Worksheet name in chart range
    ///   exceeds Excel's limit of 31 characters.
    /// - {@link XlsxError#SheetnameContainsInvalidCharacter} - Worksheet name in
    ///   chart range cannot contain invalid characters: `[ ] : * ? / \`
    /// - {@link XlsxError#SheetnameStartsOrEndsWithApostrophe} - Worksheet name in
    ///   chart range cannot start or end with an apostrophe.
    #[wasm_bindgen(js_name = "addSparklineGroup", skip_jsdoc)]
    pub fn add_sparkline_group(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        sparkline: &Sparkline,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .add_sparkline_group(
                first_row,
                first_col,
                last_row,
                last_col,
                &*sparkline.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Protect a worksheet from modification.
    ///
    /// The `protect()` method protects a worksheet from modification. It works
    /// by enabling a cell's `locked` and `hidden` properties, if they have been
    /// set. A **locked** cell cannot be edited and this property is on by
    /// default for all cells. A **hidden** cell will display the results of a
    /// formula but not the formula itself.
    ///
    /// src="https://rustxlsxwriter.github.io/images/protection_alert.png">
    ///
    /// These properties can be set using the
    /// {@link Format#setLocked}(Format::set_locked)
    /// {@link Format#setUnlocked}(Format::set_unlocked) and
    /// {@link Worksheet#setHidden}(Format::set_hidden) format methods. All cells
    /// have the `locked` property turned on by default (see the example below)
    /// so in general you don't have to explicitly turn it on.
    #[wasm_bindgen(js_name = "protect", skip_jsdoc)]
    pub fn protect(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index).unwrap().protect();
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Protect a worksheet from modification with a password.
    ///
    /// The `protect_with_password()` method is like the
    /// {@link Worksheet#protect} method, see above, except that you can add an
    /// optional, weak, password to prevent modification.
    ///
    /// **Note**: Worksheet level passwords in Excel offer very weak protection.
    /// They do not encrypt your data and are very easy to deactivate. Full
    /// workbook encryption is not supported by `rust_xlsxwriter`. However, it
    /// is possible to encrypt an `rust_xlsxwriter` file using a third party
    /// open source tool called
    /// [msoffice-crypt](https://github.com/herumi/msoffice). This works for
    /// macOS, Linux and Windows:
    ///
    /// # Parameters
    ///
    /// - `password`: The password string. Note, only ascii text passwords are
    ///   supported. Passing the empty string "" is the same as turning on
    ///   protection without a password.
    #[wasm_bindgen(js_name = "protectWithPassword", skip_jsdoc)]
    pub fn protect_with_password(&self, password: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .protect_with_password(password);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Specify which worksheet elements should, or shouldn't, be protected.
    ///
    /// The `protect_with_password()` method is like the
    /// {@link Worksheet#protect} method, see above, except it also specifies
    /// which worksheet elements should, or shouldn't, be protected.
    ///
    /// You can specify which worksheet elements protection should be on or off
    /// via a {@link ProtectionOptions} struct reference. The Excel options with
    /// their default states are shown below:
    ///
    /// src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_options1.png">
    ///
    /// # Parameters
    ///
    /// `options` - Worksheet protection options as defined by a
    /// {@link ProtectionOptions} struct reference.
    #[wasm_bindgen(js_name = "protectWithOptions", skip_jsdoc)]
    pub fn protect_with_options(&self, options: &ProtectionOptions) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .protect_with_options(&*options.inner.lock().unwrap());
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Unprotect a range of cells in a protected worksheet.
    ///
    /// As shown in the example for the {@link Worksheet#protect} method it is
    /// possible to unprotect a cell by setting the format `unprotect` property.
    /// Excel also offers an interface to unprotect larger ranges of cells. This
    /// is replicated in `rust_xlsxwriter` using the `unprotect_range()` method,
    /// see the example below.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    #[wasm_bindgen(js_name = "unprotectRange", skip_jsdoc)]
    pub fn unprotect_range(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .unprotect_range(first_row, first_col, last_row, last_col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Unprotect a range of cells in a protected worksheet, with options.
    ///
    /// This method is similar to {@link Worksheet#unprotectRange}, see above,
    /// expect that it allows you to specify two additional parameters to set
    /// the name of the range (instead of the default `Range1` .. `RangeN`) and
    /// also a optional weak password (see
    /// {@link Worksheet#protectWithPassword} for an explanation of what weak
    /// means here).
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `name`: The name of the range instead of `RangeN`. Can be blank if not
    ///   required.
    /// - `password`: The password to prevent modification of the range. Can be
    ///   blank if not required.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    #[wasm_bindgen(js_name = "unprotectRangeWithOptions", skip_jsdoc)]
    pub fn unprotect_range_with_options(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        name: &str,
        password: &str,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .unprotect_range_with_options(
                first_row, first_col, last_row, last_col, name, password,
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the selected cell or cells in a worksheet.
    ///
    /// The `set_selection()` method can be used to specify which cell or range
    /// of cells is selected in a worksheet. The most common requirement is to
    /// select a single cell, in which case the `first_` and `last_` parameters
    /// should be the same.
    ///
    /// The active cell within a selected range is determined by the order in
    /// which `first_` and `last_` are specified.
    ///
    /// Only one range of cells can be selected. The default cell selection is
    /// (0, 0, 0, 0), "A1".
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "setSelection", skip_jsdoc)]
    pub fn set_selection(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_selection(first_row, first_col, last_row, last_col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the first visible cell at the top left of a worksheet.
    ///
    /// This `set_top_left_cell()` method can be used to set the top leftmost
    /// visible cell in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "setTopLeftCell", skip_jsdoc)]
    pub fn set_top_left_cell(&self, row: u32, col: u16) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_top_left_cell(row, col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Write a user defined result to a worksheet formula cell.
    ///
    /// The `rust_xlsxwriter` library doesn’t calculate the result of a formula
    /// written using {@link Worksheet#writeFormulaWithFormat} or
    /// {@link Worksheet#writeFormula}. Instead it stores the value 0 as the
    /// formula result. It then sets a global flag in the xlsx file to say that
    /// all formulas and functions should be recalculated when the file is
    /// opened.
    ///
    /// This works fine with Excel and other spreadsheet applications. However,
    /// applications that don’t have a facility to calculate formulas will only
    /// display the 0 results.
    ///
    /// If required, it is possible to specify the calculated result of a
    /// formula using the `set_formula_result()` method.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `result`: The formula result to write to the cell.
    ///
    /// # Warnings
    ///
    /// You will get a warning if you try to set a formula result for a cell
    /// that doesn't have a formula.
    #[wasm_bindgen(js_name = "setFormulaResult", skip_jsdoc)]
    pub fn set_formula_result(&self, row: u32, col: u16, result: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_formula_result(row, col, result);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Write the default formula result for worksheet formulas.
    ///
    /// The `rust_xlsxwriter` library doesn’t calculate the result of a formula
    /// written using {@link Worksheet#writeFormulaWithFormat} or
    /// {@link Worksheet#writeFormula}. Instead it stores the value 0 as the
    /// formula result. It then sets a global flag in the xlsx file to say that
    /// all formulas and functions should be recalculated when the file is
    /// opened.
    ///
    /// However, for `LibreOffice` the default formula result should be set to
    /// the empty string literal `""`, via the `set_formula_result_default()`
    /// method, to force calculation of the result.
    ///
    /// # Parameters
    ///
    /// - `result`: The default formula result to write to the cell.
    #[wasm_bindgen(js_name = "setFormulaResultDefault", skip_jsdoc)]
    pub fn set_formula_result_default(&self, result: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_formula_result_default(result);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Add formatting to a cell without overwriting the cell data.
    ///
    /// In Excel the data in a worksheet cell is comprised of a type, a value
    /// and a format. When using `rust_xlsxwriter` the type is inferred and the
    /// value and format are generally written at the same time using methods
    /// like {@link Worksheet#writeWithFormat}. However, if required you can
    /// write the data separately and then add the format using methods like
    /// `set_cell_format()`.
    ///
    /// Although this method requires an additional step it allows for use cases
    /// where it is easier to write a large amount of data in one go and then
    /// figure out where formatting should be applied. See also the
    /// documentation section on [Worksheet Cell
    /// formatting](../worksheet/index.html#cell-formatting).
    ///
    /// For use cases where the cell formatting changes based on cell values
    /// Conditional Formatting may be a better option (see [Working with
    /// Conditional Formats](../conditional_format/index.html)).
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "setCellFormat", skip_jsdoc)]
    pub fn set_cell_format(&self, row: u32, col: u16, format: &Format) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_cell_format(row, col, &*format.inner.lock().unwrap())?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add formatting to a range of cells without overwriting the cell data.
    ///
    /// In Excel the data in a worksheet cell is comprised of a type, a value
    /// and a format. When using `rust_xlsxwriter` the type is inferred and the
    /// value and format are generally written at the same time using methods
    /// like {@link Worksheet#writeWithFormat}. However, if required you can
    /// write the data separately and then add the format using methods like
    /// `set_range_format()` or {@link Worksheet#setCellFormat} (see above).
    ///
    /// Although this method requires an additional step it allows for use cases
    /// where it is easier to write a large amount of data in one go and then
    /// figure out where formatting should be applied. See also the
    /// documentation section on [Worksheet Cell
    /// formatting](../worksheet/index.html#cell-formatting).
    ///
    /// For use cases where the cell formatting changes based on cell values
    /// Conditional Formatting may be a better option (see [Working with
    /// Conditional Formats](../conditional_format/index.html)).
    ///
    /// Note, this method should **not** be used to set the formatting for an
    /// entire worksheet since it would add an XML element for each of the 34
    /// billion cells in the worksheet which would result in a very large and
    /// slow output file. To set the format for the entire worksheet see the
    /// {@link Worksheet#setColumnRangeFormat} method.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `format`: The {@link Format} property for the cell.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    #[wasm_bindgen(js_name = "setRangeFormat", skip_jsdoc)]
    pub fn set_range_format(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_range_format(
                first_row,
                first_col,
                last_row,
                last_col,
                &*format.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Add formatting to a range of cells with an external border.
    ///
    /// This method is similar to the  {@link Worksheet#setRangeFormat} method
    /// (see above) except it also adds a border around the cell range.
    ///
    /// Add a border around a range of cells in Excel is generally easy to do
    /// using the GUI interface. However, creating a border around a range of
    /// cells programmatically is much harder since it requires the creation of
    /// up to 9 separate formats and the tracking of where cells are relative to
    /// the border.
    ///
    /// The `set_range_format_with_border()` method is provided to simplify this
    /// task. It allows you to specify one format for the cells and another for
    /// the border.
    ///
    /// You should also consider formatting a range of cells as a Worksheet
    /// Table may be a better option than simple cell formatting (see the
    /// {@link Table} section of the documentation).
    ///
    /// Note, this method should **not** be used to set the formatting for an
    /// entire worksheet since it would add an XML element for each of the 34
    /// billion cells in the worksheet which would result in a very large and
    /// slow output file. To set the format for the entire worksheet see the
    /// {@link Worksheet#setColumnRangeFormat} method.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `cell_format`: The {@link Format} property for the cells in the range. If
    ///   you don't require internal formatting you can use `Format::default()`.
    /// - `border_format`: The {@link Format} property for the border. Only the
    ///   {@link Format#setBorder} and {@link Format#setBorderColor} properties
    ///   are used.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    #[wasm_bindgen(js_name = "setRangeFormatWithBorder", skip_jsdoc)]
    pub fn set_range_format_with_border(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        cell_format: &Format,
        border_format: &Format,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_range_format_with_border(
                first_row,
                first_col,
                last_row,
                last_col,
                &*cell_format.inner.lock().unwrap(),
                &*border_format.inner.lock().unwrap(),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Clear the data and formatting from a worksheet cell.
    ///
    /// This method can be used to clear data and formatting previously written
    /// to a worksheet cell using one of the worksheet `write()` methods.
    ///
    /// This can occasionally be useful for scenarios where it is easier to add
    /// data in bulk but then remove certain elements.
    ///
    /// This method only clears data, it doesn't clear images or conditional
    /// formatting, or other non-data elements.
    ///
    /// Note, this method doesn't return a `Result` or errors. Instructions to
    /// clear non-existent cells are simply ignored.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    #[wasm_bindgen(js_name = "clearCell", skip_jsdoc)]
    pub fn clear_cell(&self, row: u32, col: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .clear_cell(row, col);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Clear the formatting from a worksheet cell.
    ///
    /// This method can be used to clear the formatting previously added to a
    /// worksheet cell using one of the worksheet `write_with_format()` methods.
    ///
    /// This can occasionally be useful for scenarios where it is easier to add
    /// formatted data in bulk but then remove the formatting from certain
    /// elements.
    ///
    /// See also the {@link Worksheet#setCellFormat} method for a similar
    /// method to change the format of a cell.
    ///
    /// Note, this method doesn't return a `Result` or errors. Instructions to
    /// clear non-existent cells are simply ignored.
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    #[wasm_bindgen(js_name = "clearCellFormat", skip_jsdoc)]
    pub fn clear_cell_format(&self, row: u32, col: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .clear_cell_format(row, col);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Display the worksheet cells from right to left for some versions of
    /// Excel.
    ///
    /// The `set_right_to_left()` method is used to change the default direction
    /// of the worksheet from left-to-right, with the A1 cell in the top left,
    /// to right-to-left, with the A1 cell in the top right.
    ///
    /// This is useful when creating Arabic, Hebrew or other near or far eastern
    /// worksheets that use right-to-left as the default direction.
    ///
    /// Depending on your use case, and text, you may also need to use the
    /// {@link Format#setReadingDirection}
    /// method to set the direction of the text within the cells.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setRightToLeft", skip_jsdoc)]
    pub fn set_right_to_left(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_right_to_left(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Make a worksheet the active/initially visible worksheet in a workbook.
    ///
    /// The `set_active()` method is used to specify which worksheet is
    /// initially visible in a multi-sheet workbook. If no worksheet is set then
    /// the first worksheet is made the active worksheet, like in Excel.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setActive", skip_jsdoc)]
    pub fn set_active(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_active(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set a worksheet tab as selected.
    ///
    /// The `set_selected()` method is used to indicate that a worksheet is
    /// selected in a multi-sheet workbook.
    ///
    /// A selected worksheet has its tab highlighted. Selecting worksheets is a
    /// way of grouping them together so that, for example, several worksheets
    /// could be printed in one go. A worksheet that has been activated via the
    /// {@link Worksheet#setActive} method will also appear as selected.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setSelected", skip_jsdoc)]
    pub fn set_selected(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_selected(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
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
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_hidden(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Hide a worksheet. Can only be unhidden in Excel by VBA.
    ///
    /// The `set_very_hidden()` method can be used to hide a worksheet similar
    /// to the {@link Worksheet#setHidden} method. The difference is that the
    /// worksheet cannot be unhidden in the the Excel user interface. The Excel
    /// worksheet `xlSheetVeryHidden` option can only be unset programmatically
    /// by VBA.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setVeryHidden", skip_jsdoc)]
    pub fn set_very_hidden(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_very_hidden(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set current worksheet as the first visible sheet tab.
    ///
    /// The {@link Worksheet#setActive}  method determines which worksheet is
    /// initially selected. However, if there are a large number of worksheets
    /// the selected worksheet may not appear on the screen. To avoid this you
    /// can select which is the leftmost visible worksheet tab using
    /// `set_first_tab()`.
    ///
    /// This method is not required very often. The default is the first
    /// worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setFirstTab", skip_jsdoc)]
    pub fn set_first_tab(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_first_tab(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the color of the worksheet tab.
    ///
    /// The `set_tab_color()` method can be used to change the color of the
    /// worksheet tab. This is useful for highlighting the important tab in a
    /// group of worksheets.
    ///
    /// # Parameters
    ///
    /// - `color`: The tab color property defined by a {@link Color} enum
    ///   value.
    #[wasm_bindgen(js_name = "setTabColor", skip_jsdoc)]
    pub fn set_tab_color(&self, color: Color) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_tab_color(xlsx::Color::from(color));
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the paper type/size when printing.
    ///
    /// This method is used to set the paper format for the printed output of a
    /// worksheet. The following paper styles are available:
    ///
    /// | Index    | Paper format            | Paper size           |
    /// | :------- | :---------------------- | :------------------- |
    /// | 0        | Printer default         | Printer default      |
    /// | 1        | Letter                  | 8 1/2 x 11 in        |
    /// | 2        | Letter Small            | 8 1/2 x 11 in        |
    /// | 3        | Tabloid                 | 11 x 17 in           |
    /// | 4        | Ledger                  | 17 x 11 in           |
    /// | 5        | Legal                   | 8 1/2 x 14 in        |
    /// | 6        | Statement               | 5 1/2 x 8 1/2 in     |
    /// | 7        | Executive               | 7 1/4 x 10 1/2 in    |
    /// | 8        | A3                      | 297 x 420 mm         |
    /// | 9        | A4                      | 210 x 297 mm         |
    /// | 10       | A4 Small                | 210 x 297 mm         |
    /// | 11       | A5                      | 148 x 210 mm         |
    /// | 12       | B4                      | 250 x 354 mm         |
    /// | 13       | B5                      | 182 x 257 mm         |
    /// | 14       | Folio                   | 8 1/2 x 13 in        |
    /// | 15       | Quarto                  | 215 x 275 mm         |
    /// | 16       | ---                     | 10x14 in             |
    /// | 17       | ---                     | 11x17 in             |
    /// | 18       | Note                    | 8 1/2 x 11 in        |
    /// | 19       | Envelope 9              | 3 7/8 x 8 7/8        |
    /// | 20       | Envelope 10             | 4 1/8 x 9 1/2        |
    /// | 21       | Envelope 11             | 4 1/2 x 10 3/8       |
    /// | 22       | Envelope 12             | 4 3/4 x 11           |
    /// | 23       | Envelope 14             | 5 x 11 1/2           |
    /// | 24       | C size sheet            | ---                  |
    /// | 25       | D size sheet            | ---                  |
    /// | 26       | E size sheet            | ---                  |
    /// | 27       | Envelope DL             | 110 x 220 mm         |
    /// | 28       | Envelope C3             | 324 x 458 mm         |
    /// | 29       | Envelope C4             | 229 x 324 mm         |
    /// | 30       | Envelope C5             | 162 x 229 mm         |
    /// | 31       | Envelope C6             | 114 x 162 mm         |
    /// | 32       | Envelope C65            | 114 x 229 mm         |
    /// | 33       | Envelope B4             | 250 x 353 mm         |
    /// | 34       | Envelope B5             | 176 x 250 mm         |
    /// | 35       | Envelope B6             | 176 x 125 mm         |
    /// | 36       | Envelope                | 110 x 230 mm         |
    /// | 37       | Monarch                 | 3.875 x 7.5 in       |
    /// | 38       | Envelope                | 3 5/8 x 6 1/2 in     |
    /// | 39       | Fanfold                 | 14 7/8 x 11 in       |
    /// | 40       | German Std Fanfold      | 8 1/2 x 12 in        |
    /// | 41       | German Legal Fanfold    | 8 1/2 x 13 in        |
    ///
    /// Note, it is likely that not all of these paper types will be available
    /// to the end user since it will depend on the paper formats that the
    /// user's printer supports. Therefore, it is best to stick to standard
    /// paper types of 1 for US Letter and 9 for A4.
    ///
    /// If you do not specify a paper type the worksheet will print using the
    /// printer's default paper style.
    ///
    /// # Parameters
    ///
    /// - `paper_size`: The paper size index from the list above .
    #[wasm_bindgen(js_name = "setPaperSize", skip_jsdoc)]
    pub fn set_paper_size(&self, paper_size: u8) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_paper_size(paper_size);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the order in which pages are printed.
    ///
    /// The `set_page_order()` method is used to change the default print
    /// direction. This is referred to by Excel as the sheet "page order":
    ///
    /// The default page order is shown below for a worksheet that extends over
    /// 4 pages. The order is called "down then over":
    ///
    /// src="https://rustxlsxwriter.github.io/images/worksheet_set_page_order.png">
    ///
    /// However, by using `set_page_order(false)` the print order will be
    /// changed to "over then down".
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. Set `true` to get "Down, then
    ///   over" (the default) and `false` to get "Over, then down".
    #[wasm_bindgen(js_name = "setPageOrder", skip_jsdoc)]
    pub fn set_page_order(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_page_order(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page orientation to landscape.
    ///
    /// The `set_landscape()` method is used to set the orientation of a
    /// worksheet's printed page to landscape.
    #[wasm_bindgen(js_name = "setLandscape", skip_jsdoc)]
    pub fn set_landscape(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_landscape();
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page orientation to portrait.
    ///
    ///  This `set_portrait()` method  is used to set the orientation of a
    ///  worksheet's printed page to portrait. The default worksheet orientation
    ///  is portrait, so this function is rarely required.
    #[wasm_bindgen(js_name = "setPortrait", skip_jsdoc)]
    pub fn set_portrait(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_portrait();
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page view mode to normal layout.
    ///
    /// This method is used to display the worksheet in “View -> Normal”
    /// mode. This is the default.
    #[wasm_bindgen(js_name = "setViewNormal", skip_jsdoc)]
    pub fn set_view_normal(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_view_normal();
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page view mode to page layout.
    ///
    /// This method is used to display the worksheet in “View -> Page Layout”
    /// mode.
    #[wasm_bindgen(js_name = "setViewPageLayout", skip_jsdoc)]
    pub fn set_view_page_layout(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_view_page_layout();
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page view mode to page break preview.
    ///
    /// This method is used to display the worksheet in “View -> Page Break
    /// Preview” mode.
    #[wasm_bindgen(js_name = "setViewPageBreakPreview", skip_jsdoc)]
    pub fn set_view_page_break_preview(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_view_page_break_preview();
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the horizontal page breaks on a worksheet.
    ///
    /// The `set_page_breaks()` method adds horizontal page breaks to a
    /// worksheet. A page break causes all the data that follows it to be
    /// printed on the next page. Horizontal page breaks act between rows.
    ///
    /// # Parameters
    ///
    /// - `breaks`: A list of one or more row numbers where the page breaks
    ///   occur. To create a page break between rows 20 and 21 you must specify
    ///   the break at row 21. However in zero index notation this is actually
    ///   row 20. So you can pretend for a small while that you are using 1
    ///   index notation.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#ParameterError} - The number of page breaks exceeds
    ///   Excel's limit of 1023 page breaks.
    #[wasm_bindgen(js_name = "setPageBreaks", skip_jsdoc)]
    pub fn set_page_breaks(&self, breaks: Vec<u32>) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_page_breaks(&breaks)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the vertical page breaks on a worksheet.
    ///
    /// The `set_vertical_page_breaks()` method adds vertical page breaks to a
    /// worksheet. This is much less common than the
    /// {@link Worksheet#setPageBreaks} method shown above.
    ///
    /// # Parameters
    ///
    /// - `breaks`: A list of one or more column numbers where the page breaks
    ///   occur.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#ParameterError} - The number of page breaks exceeds
    ///   Excel's limit of 1023 page breaks.
    #[wasm_bindgen(js_name = "setVerticalPageBreaks", skip_jsdoc)]
    pub fn set_vertical_page_breaks(&self, breaks: Vec<u32>) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_vertical_page_breaks(&breaks)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the worksheet zoom factor.
    ///
    /// Set the worksheet zoom factor in the range 10 <= zoom <= 400.
    ///
    /// The default zoom level is 100. The `set_zoom()` method does not affect
    /// the scale of the printed page in Excel. For that you should use
    /// {@link Worksheet#setPrintScale}.
    ///
    /// # Parameters
    ///
    /// - `zoom`: The worksheet zoom level.
    #[wasm_bindgen(js_name = "setZoom", skip_jsdoc)]
    pub fn set_zoom(&self, zoom: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_zoom(zoom);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set a chartsheet to automatically zoom to fit the screen.
    ///
    /// The `set_zoom_to_fit()` method only has an effect on a
    /// [Chartsheet](Worksheet::new_chartsheet) object. It ensures that the
    /// chartsheet is zoomed automatically by Excel to fit the screen even when
    /// the window is resized.
    ///
    /// This method doesn't have an effect on a standard worksheet.
    ///
    /// # Parameters
    ///
    /// - `enable`: A boolean value to enable or disable the zoom to fit.
    #[wasm_bindgen(js_name = "setZoomToFit", skip_jsdoc)]
    pub fn set_zoom_to_fit(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_zoom_to_fit(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
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
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "setFreezePanes", skip_jsdoc)]
    pub fn set_freeze_panes(&self, row: u32, col: u16) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_freeze_panes(row, col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the top most cell in the scrolling area of a freeze pane.
    ///
    /// This method is used in conjunction with the
    /// {@link Worksheet#setFreezePanes} method to set the top most visible
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
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    #[wasm_bindgen(js_name = "setFreezePanesTopCell", skip_jsdoc)]
    pub fn set_freeze_panes_top_cell(&self, row: u32, col: u16) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_freeze_panes_top_cell(row, col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the printed page header caption.
    ///
    /// The `set_header()` method can be used to set the header for a worksheet.
    ///
    /// Headers and footers are generated using a string which is a combination
    /// of plain text and optional control characters.
    ///
    /// The available control characters are:
    ///
    /// | Control              | Category      | Description           |
    /// | -------------------- | ------------- | --------------------- |
    /// | `&L`                 | Alignment     | Left                  |
    /// | `&C`                 |               | Center                |
    /// | `&R`                 |               | Right                 |
    /// | `&[Page]`  or `&P`   | Information   | Page number           |
    /// | `&[Pages]` or `&N`   |               | Total number of pages |
    /// | `&[Date]`  or `&D`   |               | Date                  |
    /// | `&[Time]`  or `&T`   |               | Time                  |
    /// | `&[File]`  or `&F`   |               | File name             |
    /// | `&[Tab]`   or `&A`   |               | Worksheet name        |
    /// | `&[Path]`  or `&Z`   |               | Workbook path         |
    /// | `&fontsize`          | Font          | Font size             |
    /// | `&"font,style"`      |               | Font name and style   |
    /// | `&U`                 |               | Single underline      |
    /// | `&E`                 |               | Double underline      |
    /// | `&S`                 |               | Strikethrough         |
    /// | `&X`                 |               | Superscript           |
    /// | `&Y`                 |               | Subscript             |
    /// | `&[Picture]` or `&G` | Images        | Picture/image         |
    /// | `&&`                 | Miscellaneous | Literal ampersand &   |
    ///
    /// Some of the placeholder variables have a long version like `&[Page]` and
    /// a short version like `&P`. The longer version is displayed in the Excel
    /// interface but the shorter version is the way that it is stored in the
    /// file format. Either version is okay since `rust_xlsxwriter` will
    /// translate as required.
    ///
    /// Headers and footers have 3 edit areas to the left, center and right.
    /// Text can be aligned to these areas by prefixing the text with the
    /// control characters `&L`, `&C` and `&R`.
    ///
    /// For example:
    ///
    /// You can also have text in each of the alignment areas:
    ///
    /// The information control characters act as variables/templates that Excel
    /// will update/expand as the workbook or worksheet changes.
    ///
    /// Times and dates are in the user's default format:
    ///
    /// To include a single literal ampersand `&` in a header or footer you
    /// should use a double ampersand `&&`:
    ///
    /// You can specify the font size of a section of the text by prefixing it
    /// with the control character `&n` where `n` is the font size:
    ///
    /// You can specify the font of a section of the text by prefixing it with
    /// the control sequence `&"font,style"` where `fontname` is a font name
    /// such as Windows font descriptions: "Regular", "Italic", "Bold" or "Bold
    /// Italic": "Courier New" or "Times New Roman" and `style` is one of the
    /// standard Windows font descriptions like “Regular”, “Italic”, “Bold” or
    /// “Bold Italic”:
    ///
    /// It is possible to combine all of these features together to create
    /// complex headers and footers. If you set up a complex header in Excel you
    /// can transfer it to `rust_xlsxwriter` by inspecting the string in the
    /// Excel file. For example the following shows how unzip and grep the Excel
    /// XML sub-files on a Linux system. The example uses libxml's xmllint to
    /// format the XML for clarity:
    ///
    /// Note: Excel requires that the header or footer string be less than 256
    /// characters, including the control characters. Strings longer than this
    /// will not be written, and a warning will be output.
    ///
    /// # Parameters
    ///
    /// - `header`: The header string with optional control characters.
    #[wasm_bindgen(js_name = "setHeader", skip_jsdoc)]
    pub fn set_header(&self, header: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_header(header);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the printed page footer caption.
    ///
    /// The `set_footer()` method can be used to set the footer for a worksheet.
    ///
    /// See the documentation for {@link Worksheet#setHeader} for more details
    /// on the syntax of the header/footer string.
    ///
    /// # Parameters
    ///
    /// - `footer`: The footer string with optional control characters.
    #[wasm_bindgen(js_name = "setFooter", skip_jsdoc)]
    pub fn set_footer(&self, footer: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_footer(footer);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Insert an image in a worksheet header.
    ///
    /// Insert an image in a worksheet header in one of the 3 sections supported
    /// by Excel: Left, Center and Right. This needs to be preceded by a call to
    /// {@link Worksheet#setHeader} where a corresponding `&[Picture]` element
    /// is added to the header formatting string such as `"&L&[Picture]"`.
    ///
    /// # Parameters
    ///
    /// - `position`: The image position as defined by the
    ///   {@link HeaderImagePosition} enum.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#ParameterError} - Parameter error if there isn't a
    ///   corresponding `&[Picture]`/`&[G]` variable in the header string.
    #[wasm_bindgen(js_name = "setHeaderImage", skip_jsdoc)]
    pub fn set_header_image(
        &self,
        image: &Image,
        position: HeaderImagePosition,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_header_image(
                &*image.inner.lock().unwrap(),
                xlsx::HeaderImagePosition::from(position),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Insert an image in a worksheet footer.
    ///
    /// See the documentation for {@link Worksheet#setHeaderImage} for more
    /// details.
    ///
    /// # Parameters
    ///
    /// - `position`: The image position as defined by the
    ///   {@link HeaderImagePosition} enum.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#ParameterError} - Parameter error if there isn't a
    ///   corresponding `&[Picture]`/`&[G]` variable in the header string.
    #[wasm_bindgen(js_name = "setFooterImage", skip_jsdoc)]
    pub fn set_footer_image(
        &self,
        image: &Image,
        position: HeaderImagePosition,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_footer_image(
                &*image.inner.lock().unwrap(),
                xlsx::HeaderImagePosition::from(position),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the page setup option to scale the header/footer with the document.
    ///
    /// This option determines whether the headers and footers use the same
    /// scaling as the worksheet. This defaults to "on" in Excel.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Header/Footer](../worksheet/index.html#page-setup---headerfooter).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "setHeaderFooterScaleWithDoc", skip_jsdoc)]
    pub fn set_header_footer_scale_with_doc(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_header_footer_scale_with_doc(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page setup option to align the header/footer with the margins.
    ///
    /// This option determines whether the headers and footers align with the
    /// left and right margins of the worksheet. This defaults to "on" in Excel.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Header/Footer](../worksheet/index.html#page-setup---headerfooter).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default. S
    #[wasm_bindgen(js_name = "setHeaderFooterAlignWithPage", skip_jsdoc)]
    pub fn set_header_footer_align_with_page(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_header_footer_align_with_page(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page margins.
    ///
    /// The `set_margins()` method is used to set the margins of the worksheet
    /// when it is printed. The units are in inches. Specifying `-1.0` for any
    /// parameter will give the default Excel value. The defaults are shown
    /// below.
    ///
    /// # Parameters
    ///
    /// - `left`: Left margin in inches. Excel default is 0.7.
    /// - `right`: Right margin in inches. Excel default is 0.7.
    /// - `top`: Top margin in inches. Excel default is 0.75.
    /// - `bottom`: Bottom margin in inches. Excel default is 0.75.
    /// - `header`: Header margin in inches. Excel default is 0.3.
    /// - `footer`: Footer margin in inches. Excel default is 0.3.
    #[wasm_bindgen(js_name = "setMargins", skip_jsdoc)]
    pub fn set_margins(
        &self,
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_margins(left, right, top, bottom, header, footer);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the first page number when printing.
    ///
    /// The `set_print_first_page_number()` method is used to set the page
    /// number of the first page when the worksheet is printed out. This option
    /// will only have and effect if you have a header/footer with the `&[Page]`
    /// control character, see {@link Worksheet#setHeader}.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Page](../worksheet/index.html#page-setup---page).
    ///
    /// # Parameters
    ///
    /// - `page_number`: The page number of the first printed page.
    #[wasm_bindgen(js_name = "setPrintFirstPageNumber", skip_jsdoc)]
    pub fn set_print_first_page_number(&self, page_number: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_print_first_page_number(page_number);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page setup option to set the print scale.
    ///
    /// Set the scale factor of the printed page, in the range 10 <= scale <=
    /// 400.
    ///
    /// The default scale factor is 100. The `set_print_scale()` method
    /// does not affect the scale of the visible page in Excel. For that you
    /// should use {@link Worksheet#setZoom}.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Page](../worksheet/index.html#page-setup---page).
    ///
    /// # Parameters
    ///
    /// - `scale`: The print scale factor.
    #[wasm_bindgen(js_name = "setPrintScale", skip_jsdoc)]
    pub fn set_print_scale(&self, scale: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_print_scale(scale);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Fit the printed area to a specific number of pages both vertically and
    /// horizontally.
    ///
    /// The `set_print_fit_to_pages()` method is used to fit the printed area to
    /// a specific number of pages both vertically and horizontally. If the
    /// printed area exceeds the specified number of pages it will be scaled
    /// down to fit. This ensures that the printed area will always appear on
    /// the specified number of pages even if the page size or margins change:
    ///
    /// The print area can be defined using the `set_print_area()` method.
    ///
    /// A common requirement is to fit the printed output to `n` pages wide but
    /// have the height be as long as necessary. To achieve this set the
    /// `height` to 0, see the example below.
    ///
    /// **Notes**:
    ///
    /// - The `set_print_fit_to_pages()` will override any manual page breaks
    ///   that are defined in the worksheet.
    ///
    /// - When using `set_print_fit_to_pages()` it may also be required to set
    ///   the printer paper size using {@link Worksheet#setPaperSize} or else
    ///   Excel will default to "US Letter".
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Page](../worksheet/index.html#page-setup---page).
    ///
    /// # Parameters
    ///
    /// - `width`: Number of pages horizontally.
    /// - `height`: Number of pages vertically.
    #[wasm_bindgen(js_name = "setPrintFitToPages", skip_jsdoc)]
    pub fn set_print_fit_to_pages(&self, width: u16, height: u16) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_print_fit_to_pages(width, height);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Center the printed page horizontally.
    ///
    /// Center the worksheet data horizontally between the margins on the
    /// printed page
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Margins](../worksheet/index.html#page-setup---margins).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintCenterHorizontally", skip_jsdoc)]
    pub fn set_print_center_horizontally(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_print_center_horizontally(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Center the printed page vertically.
    ///
    /// Center the worksheet data vertically between the margins on the printed
    /// page
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Margins](../worksheet/index.html#page-setup---margins).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintCenterVertically", skip_jsdoc)]
    pub fn set_print_center_vertically(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_print_center_vertically(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the option to turn on/off the screen gridlines.
    ///
    /// The `set_screen_gridlines()` method is use to turn on/off gridlines on
    /// displayed worksheet. It is on by default.
    ///
    /// To turn on/off the printed gridlines see the
    /// {@link Worksheet#setPrintGridlines} method below.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    #[wasm_bindgen(js_name = "setScreenGridlines", skip_jsdoc)]
    pub fn set_screen_gridlines(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_screen_gridlines(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page setup option to turn on printed gridlines.
    ///
    /// The `set_print_gridlines()` method is use to turn on/off gridlines on
    /// the printed pages. It is off by default.
    ///
    /// To turn on/off the screen gridlines see the
    /// {@link Worksheet#setScreenGridlines} method above.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintGridlines", skip_jsdoc)]
    pub fn set_print_gridlines(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_print_gridlines(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page setup option to print in black and white.
    ///
    /// This `set_print_black_and_white()` method can be used to force printing
    /// in black and white only. It is off by default.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintBlackAndWhite", skip_jsdoc)]
    pub fn set_print_black_and_white(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_print_black_and_white(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page setup option to print in draft quality.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintDraft", skip_jsdoc)]
    pub fn set_print_draft(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_print_draft(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the page setup option to print the row and column headers on the
    /// printed page.
    ///
    /// The `set_print_headings()` method turns on the row and column headers
    /// when printing a worksheet. This option is off by default.
    ///
    /// See also the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setPrintHeadings", skip_jsdoc)]
    pub fn set_print_headings(&self, enable: bool) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_print_headings(enable);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the print area for the worksheet.
    ///
    /// This method is used to specify the area of the worksheet that will be
    /// printed.
    ///
    /// In order to specify an entire row or column range such as `1:20` or
    /// `A:H` you must specify the corresponding maximum column or row range.
    /// For example:
    ///
    /// - `(0, 0, 31, 16_383) == 1:32`.
    /// - `(0, 0, 1_048_575, 12) == A:M`.
    ///
    /// In these examples 16,383 is the maximum column and 1,048,575 is the
    /// maximum row (zero indexed).
    ///
    /// See also the example below and the documentation on
    /// [Worksheet Page Setup - Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row or column is larger
    ///   than the last row or column.
    #[wasm_bindgen(js_name = "setPrintArea", skip_jsdoc)]
    pub fn set_print_area(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_print_area(first_row, first_col, last_row, last_col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the number of rows to repeat at the top of each printed page.
    ///
    /// For large Excel documents it is often desirable to have the first row or
    /// rows of the worksheet print out at the top of each page.
    ///
    /// See the example below and the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (Zero indexed.)
    /// - `last_row`: The last row of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row greater than the last
    ///   row.
    #[wasm_bindgen(js_name = "setRepeatRows", skip_jsdoc)]
    pub fn set_repeat_rows(&self, first_row: u32, last_row: u32) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_repeat_rows(first_row, last_row)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the columns to repeat at the left hand side of each printed page.
    ///
    /// For large Excel documents it is often desirable to have the first column
    /// or columns of the worksheet print out at the left hand side of each
    /// page.
    ///
    /// See the example below and the documentation on [Worksheet Page Setup -
    /// Sheet](../worksheet/index.html#page-setup---sheet).
    ///
    /// # Parameters
    ///
    /// - `first_col`: The first column of the range. (Zero indexed.)
    /// - `last_col`: The last column of the range.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row or column is larger
    ///   than the last row or column.
    #[wasm_bindgen(js_name = "setRepeatColumns", skip_jsdoc)]
    pub fn set_repeat_columns(&self, first_col: u16, last_col: u16) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_repeat_columns(first_col, last_col)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Autofit the worksheet column widths to the widest data in the column,
    /// approximately.
    ///
    /// Excel autofits columns at runtime when it has access to all of the
    /// required worksheet information as well as the Windows functions for
    /// calculating display areas based on fonts and formatting.
    ///
    /// The `rust_xlsxwriter` library doesn't have access to these Windows
    /// functions so it simulates autofit by calculating string widths based on
    /// font character metrics taken from Excel.
    ///
    /// This isn't perfect but for most cases it should be sufficient and
    /// indistinguishable from the output of Excel. However there are some
    /// limitations to be aware of when using this method:
    ///
    /// - It is based on the default Excel font type and size of Calibri 11. It
    ///   will not give accurate results for other fonts or font sizes.
    /// - It only takes formatting of numbers or dates into account if the
    ///   `enhanced_autofit` feature is enabled, which requires the {@link ssfmt}
    ///   crate. See the second example below.
    /// - Autofit is a relatively expensive operation since it performs a
    ///   calculation for all the populated cells in a worksheet. This can be
    ///   mitigated by using the {@link Worksheet#setAutofitMaxRow} method,
    ///   see below.
    ///
    /// For cases that don't match your desired output you can set explicit
    /// column widths via {@link Worksheet#setColumnWidth} or
    /// {@link Worksheet#setColumnWidthPixels}. The `autofit()` method ignores
    /// columns that have already been explicitly set if the width is greater
    /// than the calculated autofit width. Alternatively, setting the column
    /// width explicitly after calling `autofit()` will override the autofit
    /// value.
    #[wasm_bindgen(js_name = "autofit", skip_jsdoc)]
    pub fn autofit(&self) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index).unwrap().autofit();
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the maximum row used for autofitting worksheet columns.
    ///
    /// The {@link Worksheet#autofit} method simulates Excel's column autofit by
    /// calculating the width of each cell based on its contents formatted as a
    /// string which, for large datasets, this can be an expensive operation. In
    /// order to mitigate this the `set_autofit_max_row()` method can be used to
    /// limit the number of rows processed for autofitting. This takes advantage
    /// of the fact that a user will typically only see about 50 to 100 rows on
    /// a screen so it is often sufficient to autofit just the first few hundred
    /// rows. This can significantly speed up the autofit operation for large
    /// datasets.
    ///
    /// A value of 200 rows is recommended as a good compromise between
    /// performance and visual accuracy.
    ///
    /// # Parameters
    ///
    /// - `max_row`: The maximum row number to use for autofitting.
    #[wasm_bindgen(js_name = "setAutofitMaxRow", skip_jsdoc)]
    pub fn set_autofit_max_row(&self, max_row: u32) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_autofit_max_row(max_row);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the maximum autofit width for worksheet columns.
    ///
    /// The {@link Worksheet#autofit} method simulates Excel's column autofit.
    /// One undesirable side-effect of this is that Excel autofits very long
    /// strings up to limit of 255 characters/1790 pixels. This is often too
    /// wide to display on a single screen at normal zoom. As such the
    /// `set_autofit_max_width()` method can be used to set a smaller upper
    /// limit for autofitting long strings.
    ///
    /// A value of 300 pixels is recommended as a good compromise between column
    /// width and readability.
    ///
    /// # Parameters
    ///
    /// - `max_width`: The maximum column width, in pixels, to use for autofitting.
    #[wasm_bindgen(js_name = "setAutofitMaxWidth", skip_jsdoc)]
    pub fn set_autofit_max_width(&self, max_width: u32) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_autofit_max_width(max_width);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Autofit the worksheet columns up to a maximum width.
    ///
    /// The {@link Worksheet#autofit} method above simulates Excel's column
    /// autofit. One undesirable side-effect of this is that Excel autofits very
    /// long strings up to limit of 255 characters/1790 pixels. This is often
    /// too wide to display on a single screen at normal zoom. As such the
    /// `autofit_to_max_width()` method is provided to enable a smaller upper
    /// limit for autofitting long strings. A value of 300 pixels is recommended
    /// as a good compromise between column width and readability.
    ///
    /// # Parameters
    ///
    /// - `max_width`: The maximum column width, in pixels, to use for
    ///   autofitting.
    #[wasm_bindgen(js_name = "autofitToMaxWidth", skip_jsdoc)]
    pub fn autofit_to_max_width(&self, max_width: u32) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .autofit_to_max_width(max_width);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the worksheet name used in VBA macros.
    ///
    /// This method can be used to set the VBA name for the worksheet. This is
    /// sometimes required when a VBA macro included via
    /// {@link Workbook#addVbaProject})
    /// makes reference to the worksheet with a name other than the default
    /// Excel VBA names of `Sheet1`, `Sheet2`, etc.
    ///
    /// See also the
    /// {@link Workbook#setVbaName}) method for
    /// setting the workbook VBA name.
    ///
    /// The name must be a valid Excel VBA object name as defined by the
    /// following rules:
    ///
    /// - The name must be less than 32 characters.
    /// - The name can only contain word characters: letters, numbers and
    ///   underscores.
    /// - The name must start with a letter.
    /// - The name cannot be blank.
    ///
    /// The name must be also be unique across the worksheets in the workbook.
    ///
    /// # Parameters
    ///
    /// - `name`: The vba name. It must follow the Excel rules, shown above.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#VbaNameError} - The name doesn't meet one of Excel's
    ///   criteria, shown above.
    #[wasm_bindgen(js_name = "setVbaName", skip_jsdoc)]
    pub fn set_vba_name(&self, name: &str) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_vba_name(name)?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Set the default string used for NaN values.
    ///
    /// Excel doesn't support storing `NaN` (Not a Number) values. If a `NAN` is
    /// generated as the result of a calculation Excel stores and displays the
    /// error `#NUM!`. However, this error isn't usually used outside of a
    /// formula result and it isn't stored as a number.
    ///
    /// In order to deal with `f64::NAN` in a reasonable way `rust_xlsxwriter`
    /// writes it as the string "NAN". The `set_nan_value()` method allows you
    /// to override this default value.
    ///
    /// # Parameters
    ///
    /// - `value`: The string to use for NaN values.
    #[wasm_bindgen(js_name = "setNanValue", skip_jsdoc)]
    pub fn set_nan_value(&self, value: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_nan_value(value);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the default string used for Infinite values.
    ///
    /// Excel doesn't support storing `Inf` (infinity) values. If an `Inf` is
    /// generated as the result of a calculation Excel stores and displays the
    /// error `#DIV/0`. However, this error isn't usually used outside of a
    /// formula result and it isn't stored as a number.
    ///
    /// In order to deal with `f64::INFINITY` in a reasonable way
    /// `rust_xlsxwriter` writes it as the string "INF". The
    /// `set_infinite_value()` method allows you to override this default value.
    ///
    /// See the example for {@link Worksheet#setNanValue} above.
    ///
    /// # Parameters
    ///
    /// - `value`: The string to use for `Inf` values.
    #[wasm_bindgen(js_name = "setInfinityValue", skip_jsdoc)]
    pub fn set_infinity_value(&self, value: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_infinity_value(value);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Set the default string used for negative Infinite values.
    ///
    /// Excel doesn't support storing `-Inf` (negative infinity) values. If a
    /// `-Inf` is generated as the result of a calculation Excel stores and
    /// displays the error `#DIV/0`. However, this error isn't usually used
    /// outside of a formula result and it isn't stored as a number.
    ///
    /// In order to deal with `f64::NEG_INFINITY` in a reasonable way
    /// `rust_xlsxwriter` writes it as the string "-INF". The
    /// `set_infinite_value()` method method allows you to override this default
    /// value.
    ///
    /// See the example for {@link Worksheet#setNanValue} above.
    ///
    /// # Parameters
    ///
    /// - `value`: The string to use for `-Inf` values.
    #[wasm_bindgen(js_name = "setNegInfinityValue", skip_jsdoc)]
    pub fn set_neg_infinity_value(&self, value: &str) -> Worksheet {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .set_neg_infinity_value(value);
        Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        }
    }
    /// Ignore an Excel error or warning in a worksheet cell.
    ///
    /// Excel flags a number of data errors and inconsistencies with a a small
    /// green triangle in the top left hand corner of the cell. For example the
    /// following causes a warning of "Number Stored as Text":
    ///
    /// src="https://rustxlsxwriter.github.io/images/worksheet_ignore_error1.png">
    ///
    /// These warnings can be useful indicators that there is an issue in the
    /// spreadsheet but sometimes it is preferable to turn them off. At the file
    /// level these errors can be ignored for a cell or cell range using
    /// `Worksheet::ignore_error()` and {@link Worksheet#ignoreErrorRange} (see
    /// below).
    ///
    /// The errors and warnings that can be turned off at the file level are
    /// represented by the {@link IgnoreError} enum values. These equate, with some
    /// minor exceptions, to the error categories shown in the Excel Error
    /// Checking dialog:
    ///
    /// src="https://rustxlsxwriter.github.io/images/ignore_errors_dialog.png">
    ///
    /// (Note: some of the items shown in the above dialog such as "Misleading
    /// Number Formats" aren't saved in the output file by Excel and can't be
    /// turned off permanently.)
    ///
    /// <br>
    ///
    /// The `Worksheet::ignore_error()` method can be called repeatedly to
    /// ignore errors in different cells but **Excel only allows one ignored
    /// error per cell**.
    ///
    /// An error can be turned off for a range of cells using the
    /// {@link Worksheet#ignoreErrorRange} method (see below).
    ///
    /// # Parameters
    ///
    /// - `row`: The zero indexed row number.
    /// - `col`: The zero indexed column number.
    /// - `error_type`: An {@link IgnoreError} enum value.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#ParameterError} - Parameter error if more than one rule
    ///   is added to the same cell.
    #[wasm_bindgen(js_name = "ignoreError", skip_jsdoc)]
    pub fn ignore_error(
        &self,
        row: u32,
        col: u16,
        error_type: IgnoreError,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .ignore_error(row, col, xlsx::IgnoreError::from(error_type))?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Ignore an Excel error or warning in a range of worksheet cells.
    ///
    /// See {@link Worksheet#ignoreError} above for a detailed explanation of
    /// Excel worksheet errors.
    ///
    /// The `Worksheet::ignore_error_range()` method can be used to ignore an
    /// error in a range, a row, a column, or the entire worksheet and it can be
    /// called repeatedly to ignore errors in different cells ranges. It should
    /// be noted however that **Excel only allows one ignored error per cell**.
    /// The `rust_xlsxwriter` library verifies that multiple rules aren't added
    /// to the same cell or cell range and will raise an error if this occurs.
    /// However it doesn't currently verify whether cells within ranges overlap.
    /// It is up to the user to ensure that this doesn't happen when using
    /// ranges.
    ///
    /// # Parameters
    ///
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    /// - `error_type`: An {@link IgnoreError} enum value.
    ///
    /// # Errors
    ///
    /// - {@link XlsxError#RowColumnLimitError} - Row or column exceeds Excel's
    ///   worksheet limits.
    /// - {@link XlsxError#RowColumnOrderError} - First row or column is larger
    ///   than the last row or column.
    /// - {@link XlsxError#ParameterError} - Parameter error if more than one rule
    ///   is added to the same range.
    #[wasm_bindgen(js_name = "ignoreErrorRange", skip_jsdoc)]
    pub fn ignore_error_range(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        error_type: IgnoreError,
    ) -> WasmResult<Worksheet> {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .ignore_error_range(
                first_row,
                first_col,
                last_row,
                last_col,
                xlsx::IgnoreError::from(error_type),
            )?;
        Ok(Worksheet {
            parent: Arc::clone(&self.parent),
            index: self.index,
        })
    }
    /// Get the local instance DXF id for a format.
    ///
    /// Get the local instance DXF id for a format. These indexes will be
    /// replaced by global/workbook indices before the worksheet is saved. DXF
    /// indexed are used for Tables and Conditional Formats.
    ///
    /// This method is public but hidden to allow test cases to mirror the
    /// creation order for DXF ids which is usually the reverse of the order of
    /// the XF instance ids.
    ///
    /// # Parameters
    ///
    /// `format` - The {@link Format} instance to register.
    #[wasm_bindgen(js_name = "formatDxfIndex", skip_jsdoc)]
    pub fn format_dxf_index(&self, format: &Format) -> u32 {
        let mut lock = self.parent.lock().unwrap();
        lock.worksheet_from_index(self.index)
            .unwrap()
            .format_dxf_index(&*format.inner.lock().unwrap())
    }
}
