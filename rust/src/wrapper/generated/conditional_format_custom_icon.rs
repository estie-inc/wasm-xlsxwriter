use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatCustomIcon {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatCustomIcon>>,
}

#[wasm_bindgen]
impl ConditionalFormatCustomIcon {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatCustomIcon {
        ConditionalFormatCustomIcon {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatCustomIcon::new())),
        }
    }
}
