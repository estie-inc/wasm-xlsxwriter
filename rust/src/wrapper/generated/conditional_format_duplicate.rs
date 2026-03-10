use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatDuplicate` struct represents a Duplicate/Unique
/// conditional format.
///
/// `ConditionalFormatDuplicate` is used to represent a Duplicate or Unique
/// style conditional format in Excel. Duplicate conditional formats show
/// duplicated values in a range while Unique conditional formats show unique
/// values.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_duplicate_intro.png">
///
/// For more information see Working with Conditional Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatDuplicate {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatDuplicate>>,
}

#[wasm_bindgen]
impl ConditionalFormatDuplicate {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatDuplicate {
        ConditionalFormatDuplicate {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatDuplicate::new())),
        }
    }
}
