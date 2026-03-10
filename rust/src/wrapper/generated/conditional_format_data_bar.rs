use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatDataBar` struct represents a Data Bar conditional
/// format.
///
/// `ConditionalFormatDataBar` is used to represent a Cell style conditional
/// format in Excel. A Data Bar Cell conditional format shows a per cell color
/// gradient from the minimum value to the maximum value.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_databar_intro.png">
///
/// The options that can be applied to a data bar conditional format in Excel
/// are shown in the image below.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_databar_intro2.png">
///
/// The methods to replicate these options are explained in the following
/// documentation.
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatDataBar {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatDataBar>>,
}

#[wasm_bindgen]
impl ConditionalFormatDataBar {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatDataBar {
        ConditionalFormatDataBar {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatDataBar::new())),
        }
    }
}
