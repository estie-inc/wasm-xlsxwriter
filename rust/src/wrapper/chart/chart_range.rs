// Companion: ChartRange custom constructors (upstream has no new())

use std::sync::{Arc, Mutex};

use rust_xlsxwriter::{self as xlsx, IntoChartRange};
use wasm_bindgen::prelude::*;

use crate::wrapper::ChartRange;

#[wasm_bindgen]
impl ChartRange {
    /// Create a new `ChartRange` from an Excel range formula such as `"Sheet1!$A$1:$A$3"`.
    #[wasm_bindgen(js_name = "newFromString")]
    pub fn new_from_string(range_str: &str) -> ChartRange {
        let range: xlsx::ChartRange = range_str.new_chart_range();
        ChartRange {
            inner: Arc::new(Mutex::new(range)),
        }
    }

    /// Create a new `ChartRange` from a worksheet 5 tuple.
    #[wasm_bindgen(js_name = "newFromRange")]
    pub fn new_from_range(
        sheet: &str,
        first_row: xlsx::RowNum,
        first_col: xlsx::ColNum,
        last_row: xlsx::RowNum,
        last_col: xlsx::ColNum,
    ) -> ChartRange {
        let range = (sheet, first_row, first_col, last_row, last_col).new_chart_range();
        ChartRange {
            inner: Arc::new(Mutex::new(range)),
        }
    }
}
