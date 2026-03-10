use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormatIconSet` struct represents an Icon Set style
/// conditional format.
///
/// `ConditionalFormatIconSet` is used to represent an Icon Set style
/// conditional format in Excel. An Icon Set conditional format highlights items
/// with groups of 3-5 symbols such as traffic lights, arrows, or flags.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_icon_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatIconSet {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatIconSet>>,
}

#[wasm_bindgen]
impl ConditionalFormatIconSet {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatIconSet {
        ConditionalFormatIconSet {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatIconSet::new())),
        }
    }
}
