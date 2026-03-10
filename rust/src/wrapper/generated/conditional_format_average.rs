use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

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
