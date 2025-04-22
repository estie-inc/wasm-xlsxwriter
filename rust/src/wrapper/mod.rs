mod chart;
mod color;
mod doc_properties;
mod excel_data;
mod format;
mod formula;
mod header_image_position;
mod image;
mod note;
mod rich_string;
mod table;
mod url;
mod utils;
mod workbook;
mod worksheet;

use crate::error::XlsxError;
use wasm_bindgen::prelude::wasm_bindgen;


type WasmResult<T> = std::result::Result<T, XlsxError>;

// This runs once when the wasm module is instantiated
// https://rustwasm.github.io/wasm-bindgen/reference/attributes/on-rust-exports/start.html
#[wasm_bindgen(start)]
pub fn start() {
    console_error_panic_hook::set_once();
}
