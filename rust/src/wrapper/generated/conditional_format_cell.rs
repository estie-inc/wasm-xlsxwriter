use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatCell {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatCell>>,
}

#[wasm_bindgen]
impl ConditionalFormatCell {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatCell {
        ConditionalFormatCell {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatCell::new())),
        }
    }
}
