use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ProtectionOptions {
    pub(crate) inner: Arc<Mutex<xlsx::ProtectionOptions>>,
}

#[wasm_bindgen]
impl ProtectionOptions {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ProtectionOptions {
        ProtectionOptions {
            inner: Arc::new(Mutex::new(xlsx::ProtectionOptions::new())),
        }
    }
}
