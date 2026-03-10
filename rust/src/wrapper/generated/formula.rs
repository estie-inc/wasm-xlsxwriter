use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Formula {
    pub(crate) inner: Arc<Mutex<xlsx::Formula>>,
}

#[wasm_bindgen]
impl Formula {
    #[wasm_bindgen(constructor)]
    pub fn new(formula: impl AsRef) -> Formula {
        Formula {
            inner: Arc::new(Mutex::new(xlsx::Formula::new(formula))),
        }
    }
}
