use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatFormula` struct represents a Formula style conditional
/// format.
///
/// `ConditionalFormatFormula` is used to represent a Formula style conditional
/// format in Excel. A Formula conditional format highlights formula values in a
/// range based on a user supplied formula.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_formula_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatFormula {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatFormula>>,
}

#[wasm_bindgen]
impl ConditionalFormatFormula {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatFormula {
        ConditionalFormatFormula {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatFormula::new())),
        }
    }
}
