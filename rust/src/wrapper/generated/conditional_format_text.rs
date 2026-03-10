use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatText {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatText>>,
}

#[wasm_bindgen]
impl ConditionalFormatText {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatText {
        ConditionalFormatText {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatText::new())),
        }
    }
}
