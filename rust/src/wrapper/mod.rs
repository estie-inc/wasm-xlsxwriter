mod chart;
mod conditional_format_cell;
mod data_validation;
mod datetime;
mod filter_condition;
mod excel_data;
pub mod generated;
mod image;
mod rich_string;
mod shape;
mod utils;
mod workbook;
mod worksheet;

use crate::error::XlsxError;
use wasm_bindgen::prelude::wasm_bindgen;


pub(crate) type WasmResult<T> = std::result::Result<T, XlsxError>;

// Generated types are re-exported via `generated::*`
pub(crate) use generated::*;
// Hand-written types that can't be auto-generated
pub(crate) use datetime::ExcelDateTime;
pub(crate) use workbook::Workbook;

// This runs once when the wasm module is instantiated
// https://rustwasm.github.io/wasm-bindgen/reference/attributes/on-rust-exports/start.html
#[wasm_bindgen(start)]
pub fn start() {
    console_error_panic_hook::set_once();
}
