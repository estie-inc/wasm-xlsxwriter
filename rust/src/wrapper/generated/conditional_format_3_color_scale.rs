use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ConditionalFormat3ColorScale` struct represents a 3 Color Scale
/// conditional format.
///
/// `ConditionalFormat3ColorScale` is used to represent a Cell style conditional
/// format in Excel. A 3 Color Scale Cell conditional format shows a per cell
/// color gradient from the minimum value to the maximum value.
///
/// src="https://rustxlsxwriter.github.io/images/conditional_format_3color_intro.png">
///
/// For more information see Working with Conditional
/// Formats.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormat3ColorScale {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormat3ColorScale>>,
}

#[wasm_bindgen]
impl ConditionalFormat3ColorScale {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormat3ColorScale {
        ConditionalFormat3ColorScale {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormat3ColorScale::new())),
        }
    }
}
