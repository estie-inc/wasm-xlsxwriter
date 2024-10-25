use rust_xlsxwriter::{self as xlsx, IntoChartRange};
use wasm_bindgen::prelude::*;

/// Trait to map types into an `ChartRange`.
///
/// The 2 most common types of range used in `rust_xlsxwriter` charts are:
///
/// - A string with an Excel like range formula such as `"Sheet1!$A$1:$A$3"`.
/// - A 5 value tuple that can be used to create the range programmatically
///   using a sheet name and zero indexed row and column values like:
///   `("Sheet1", 0, 0, 2, 0)` (this gives the same range as the previous string
///   value).
///
/// For single cell ranges used in chart items such as chart or axis titles you
/// can also use:
///
/// - A simple string title.
/// - A string with an Excel like cell formula such as `"Sheet1!$A$1"`.
/// - A 3 value tuple that can be used to create the cell range programmatically
///   using a sheet name and zero indexed row and column values like:
///   `("Sheet1", 0, 0)` (this gives the same range as the previous string
///   value).
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartRange {
    pub(crate) inner: xlsx::ChartRange,
}

#[wasm_bindgen]
impl ChartRange {
    /// Create a new `ChartRange` from a worksheet 5 tuple.
    ///
    /// A 5 value tuple that can be used to create the range programmatically
    ///   using a sheet name and zero indexed row and column values like:
    ///   `("Sheet1", 0, 0, 2, 0)` (this gives the same range as the previous string
    ///   value).
    #[wasm_bindgen(constructor)]
    pub fn new(
        sheet: &str,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
    ) -> ChartRange {
        let range = (sheet, first_row, first_col, last_row, last_col).new_chart_range();
        ChartRange { inner: range }
    }

    /// Create a new `ChartRange` from an Excel range formula such as `"Sheet1!$A$1:$A$3"`.
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "newFromString")]
    pub fn new_from_string(range_str: &str) -> ChartRange {
        let range: xlsx::ChartRange = range_str.new_chart_range();
        ChartRange { inner: range }
    }

    /// Create a new `ChartRange` from a worksheet 3 tuple.
    ///
    /// A 3 value tuple that can be used to create the cell range programmatically
    ///   using a sheet name and zero indexed row and column values like:
    ///   `("Sheet1", 0, 0)` (this gives the same range as the previous string
    ///   value).
    ///
    /// TODO: example omitted
    #[wasm_bindgen(js_name = "fromCell")]
    pub fn from_cell(sheet: &str, row: xlsx::RowNum, col: xlsx::ColNum) -> ChartRange {
        let range = (sheet, row, col).new_chart_range();
        ChartRange { inner: range }
    }

    // TODO: support (str, RowNum, ColNum) and (str, str) constructors
}
