use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatDate {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatDate>>,
}

#[wasm_bindgen]
impl ConditionalFormatDate {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatDate {
        ConditionalFormatDate {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatDate::new())),
        }
    }
}
