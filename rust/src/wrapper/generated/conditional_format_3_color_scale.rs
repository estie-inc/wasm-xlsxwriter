use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormat3ColorScale {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormat3ColorScale>>,
}

#[wasm_bindgen]
impl ConditionalFormat3ColorScale {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormat3ColorScale {
        ConditionalFormat3ColorScale {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormat3ColorScale::new())),
        }
    }
}
