use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `FilterData` struct represents data types used in Excel's filters.
///
/// The `FilterData` struct is a simple data type to allow a generic mapping
/// between Rust's string and number types and similar types used in Excel's
/// filters.
#[derive(Clone)]
#[wasm_bindgen]
pub struct FilterData {
    pub(crate) inner: Arc<Mutex<xlsx::FilterData>>,
}

#[wasm_bindgen]
impl FilterData {}
