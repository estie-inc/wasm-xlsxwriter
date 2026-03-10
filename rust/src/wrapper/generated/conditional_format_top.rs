use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatTop` struct represents a Top/Bottom style conditional
/// format.
///
/// `ConditionalFormatTop` is used to represent a Top or Bottom style
/// conditional format in Excel. Top conditional formats show the top X values
/// in a range. The value of the conditional can be a rank, i.e., Top X, or a
/// percentage, i.e., Top X%.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_top_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatTop {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatTop>>,
}

#[wasm_bindgen]
impl ConditionalFormatTop {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatTop {
        ConditionalFormatTop {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatTop::new())),
        }
    }
}
