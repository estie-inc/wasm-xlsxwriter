use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatError {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatError>>,
}

#[wasm_bindgen]
impl ConditionalFormatError {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatError {
        ConditionalFormatError {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatError::new())),
        }
    }
}
