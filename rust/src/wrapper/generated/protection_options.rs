use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ProtectionOptions` struct is used to set protected elements in a worksheet.
///
/// You can specify which worksheet elements should be protected or unprotected via
/// the `ProtectionOptions` members. The corresponding Excel options with
/// their default states are shown below:
///
/// src="https://rustxlsxwriter.github.io/images/worksheet_protect_with_options1.png">
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
