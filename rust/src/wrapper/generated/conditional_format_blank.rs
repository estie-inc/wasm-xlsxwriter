use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatBlank` struct represents a Blank/Non-blank conditional
/// format.
///
/// `ConditionalFormatBlank` is used to represent a Blank or Non-blank style
/// conditional format in Excel. A Blank conditional format highlights blank
/// values in a range while the inverted version highlights non-blanks values.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_blank_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatBlank {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatBlank>>,
}

#[wasm_bindgen]
impl ConditionalFormatBlank {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatBlank {
        ConditionalFormatBlank {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatBlank::new())),
        }
    }
}
