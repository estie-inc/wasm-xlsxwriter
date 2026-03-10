use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatDate` struct represents a Dates Occurring style
/// conditional format.
///
/// `ConditionalFormatDate` is used to represent a Dates Occurring style
/// conditional format in Excel. This is used to identify dates in ranges like
/// "Last Week" or "Last Month".
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_date_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatDate {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatDate>>,
}

#[wasm_bindgen]
impl ConditionalFormatDate {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatDate {
        ConditionalFormatDate {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatDate::new())),
        }
    }
}
