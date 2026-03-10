use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatCell` struct represents a Cell conditional format.
///
/// `ConditionalFormatCell` is used to represent a Cell style conditional format
/// in Excel. Cell conditional formats use simple equalities such as "equal to"
/// or "greater than" or "between".
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_cell_intro.png">
///
/// For more information see Working with Conditional Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatCell {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatCell>>,
}

#[wasm_bindgen]
impl ConditionalFormatCell {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatCell {
        ConditionalFormatCell {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatCell::new())),
        }
    }
}
