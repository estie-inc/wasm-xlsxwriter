use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatAverage` struct represents an Average/Standard
/// Deviation style conditional format.
///
/// `ConditionalFormatAverage` is used to represent a Average or Standard Deviation style
/// conditional format in Excel.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_average_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatAverage {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatAverage>>,
}

#[wasm_bindgen]
impl ConditionalFormatAverage {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatAverage {
        ConditionalFormatAverage {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatAverage::new())),
        }
    }
}
