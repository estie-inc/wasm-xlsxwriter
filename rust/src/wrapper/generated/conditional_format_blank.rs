use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatBlank {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatBlank>>,
}

#[wasm_bindgen]
impl ConditionalFormatBlank {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatBlank {
        ConditionalFormatBlank {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatBlank::new())),
        }
    }
}
