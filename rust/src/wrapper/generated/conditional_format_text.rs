use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatText` struct represents a Text conditional format.
///
/// `ConditionalFormatText` is used to represent a Text style conditional format
/// in Excel. Text conditional formats use simple equalities such as "equal to"
/// or "greater than" or "between".
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_text_intro.png">
///
/// For more information see Working with Conditional Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatText {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatText>>,
}

#[wasm_bindgen]
impl ConditionalFormatText {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatText {
        ConditionalFormatText {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatText::new())),
        }
    }
}
