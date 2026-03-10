use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

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
