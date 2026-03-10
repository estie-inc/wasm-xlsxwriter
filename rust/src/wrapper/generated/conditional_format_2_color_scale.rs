use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormat2ColorScale` struct represents a 2 Color Scale
/// conditional format.
///
/// `ConditionalFormat2ColorScale` is used to represent a Cell style conditional
/// format in Excel. A 2 Color Scale Cell conditional format shows a per cell
/// color gradient from the minimum value to the maximum value.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_2color_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormat2ColorScale {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormat2ColorScale>>,
}

#[wasm_bindgen]
impl ConditionalFormat2ColorScale {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormat2ColorScale {
        ConditionalFormat2ColorScale {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormat2ColorScale::new())),
        }
    }
}
