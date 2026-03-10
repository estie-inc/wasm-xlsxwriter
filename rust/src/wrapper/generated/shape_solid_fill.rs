use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeSolidFill {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeSolidFill>>,
}

#[wasm_bindgen]
impl ShapeSolidFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeSolidFill {
        ShapeSolidFill {
            inner: Arc::new(Mutex::new(xlsx::ShapeSolidFill::new())),
        }
    }
}
