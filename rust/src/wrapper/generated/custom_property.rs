use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `CustomProperty` struct represents data types used in Excel’s custom
/// document properties.
#[derive(Clone)]
#[wasm_bindgen]
pub struct CustomProperty {
    pub(crate) inner: Arc<Mutex<xlsx::CustomProperty>>,
}

#[wasm_bindgen]
impl CustomProperty {
    #[doc = r" Create a deep clone of this object."]
    #[wasm_bindgen(js_name = "clone")]
    pub fn deep_clone(&self) -> CustomProperty {
        CustomProperty {
            inner: Arc::new(Mutex::new(self.inner.lock().unwrap().clone())),
        }
    }
}
