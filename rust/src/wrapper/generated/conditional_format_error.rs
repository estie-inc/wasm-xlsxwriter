use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatError` struct represents an Error/Non-error conditional
/// format.
///
/// `ConditionalFormatError` is used to represent an Error or Non-error style
/// conditional format in Excel. An error conditional format highlights error
/// values in a range while the inverted version highlights non-errors values.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_error_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatError {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatError>>,
}

#[wasm_bindgen]
impl ConditionalFormatError {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatError {
        ConditionalFormatError {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatError::new())),
        }
    }
}
