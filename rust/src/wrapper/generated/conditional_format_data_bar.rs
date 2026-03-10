use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

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
