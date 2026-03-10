use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatDuplicate {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatDuplicate>>,
}

#[wasm_bindgen]
impl ConditionalFormatDuplicate {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatDuplicate {
        ConditionalFormatDuplicate {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatDuplicate::new())),
        }
    }
}
