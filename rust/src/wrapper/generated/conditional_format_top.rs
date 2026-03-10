use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatTop {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatTop>>,
}

#[wasm_bindgen]
impl ConditionalFormatTop {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatTop {
        ConditionalFormatTop {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatTop::new())),
        }
    }
}
